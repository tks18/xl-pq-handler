# parser.py
import re
from typing import Any, List, Dict, Tuple

# --- REGEX DEFINITIONS FOR M-CODE ---

# M-Code keywords
M_KEYWORDS = (
    'let', 'in', 'if', 'then', 'else', 'each', 'try', 'otherwise',
    'type', 'meta', 'as', 'is', 'section', 'shared', 'true', 'false', 'null'
)

M_DATA_SOURCES = (
    'Sql.Database', 'Web.Contents', 'File.Contents', 'Excel.Workbook',
    'Excel.CurrentWorkbook', 'Csv.Document', 'Json.Document', 'Odbc.DataSource',
    'Folder.Files', 'SharePoint.Files', 'PowerBI.Dataflows'
)

# We compile them into a single regex for performance.
TOKEN_SPEC = [
    ('COMMENT', r'//[^\n]*|/\*.*?\*/'),
    ('STRING', r'#"[^"]*"|"[^"]*"'),
    ('KEYWORD', r'\b(' + '|'.join(M_KEYWORDS) + r')\b'),
    # We must check for specific Data Sources FIRST.
    ('DATASOURCE', r'\b(' + '|'.join(re.escape(f)
     for f in M_DATA_SOURCES) + r')\b'),
    # NOW we add your generic rule for *any other* Library.Function
    # This matches 'Table.AddColumn', 'List.Max', 'Text.Proper', etc.
    ('FUNCTION', r'\b[A-Z][a-zA-Z0-9_]*\.[A-Z][a-zA-Z0-9_.]*\b'),
    ('NUMBER', r'\b\d+(\.\d*)?\b'),
    ('OPERATOR', r'[=<>\[\]{}()&+-/*,;]'),
    ('IDENTIFIER', r'\b[a-zA-Z_][a-zA-Z0-9_]*\b'),  # User-defined vars
    ('OTHER', r'.'),  # Any other character
]

# The master regex
TOKEN_REGEX = re.compile(
    '|'.join(f'(?P<{name}>{pattern})' for name, pattern in TOKEN_SPEC),
    re.DOTALL | re.IGNORECASE
)

# --- PUBLIC PARSER FUNCTIONS ---


class ExcelMCodeParser:
    """
    Set of Functions to parse M-Code.
    """

    def _parse_source_argument(self, arg: str) -> Dict[str, Any]:
        """
        Analyzes a captured source argument and classifies it.
        Returns: {"value": "...", "type": "..."}
        """
        arg = arg.strip()

        # 1. Check for literal strings
        if arg.startswith('"') and arg.endswith('"'):
            return {"value": arg[1:-1], "type": "Literal"}
        if arg.startswith('#"') and arg.endswith('"'):
            return {"value": arg[2:-1], "type": "Literal"}

        # 2. Check for records or lists
        if arg.startswith('['):
            return {"value": "[Record]", "type": "Record"}
        if arg.startswith('{'):
            return {"value": "{List}", "type": "List"}

        # 3. Check if it's another function call
        if '(' in arg and arg.endswith(')'):
            return {"value": arg, "type": "Other Function"}

        # 4. Otherwise, it's a variable or parameter
        return {"value": arg, "type": "Variable"}

    def _find_ultimate_source(self, arg_str: str) -> Dict[str, str]:
        """
        Analyzes an argument string to find the *ultimate* source.
        e.g., "File.Contents(filePath)" -> {"type": "Variable", "value": "filePath"}
        e.g., "Web.Contents("http...")" -> {"type": "Literal", "value": "http..."}
        """
        arg_str = arg_str.strip()

        # 1. Base Case: Is it a literal string?
        if (arg_str.startswith('"') and arg_str.endswith('"')):
            return {"type": "Literal", "value": arg_str[1:-1]}
        if (arg_str.startswith('#"') and arg_str.endswith('"')):
            return {"type": "Literal", "value": arg_str[2:-1]}

        # 2. Base Case: Is it just a variable?
        if re.fullmatch(r'[a-zA-Z_][a-zA-Z0-9_]*', arg_str):
            return {"type": "Variable", "value": arg_str}

        # 3. Recursive Case: Is it another function call?
        match = re.match(r'([a-zA-Z0-9_.]+)\s*\((.*)\)', arg_str, re.DOTALL)
        if match:
            # It's a function call, like File.Contents(filePath)
            inner_args_str = match.group(2)
            if re.fullmatch(r'[a-zA-Z_][a-zA-Z0-9_]*', inner_args_str.strip()):
                return {"type": "Variable", "value": inner_args_str.strip()}

            # If the nested argument is a literal, it's a literal
            if (inner_args_str.strip().startswith('"')):
                return {"type": "Literal", "value": inner_args_str.strip(' "')}

        # 4. Fallback: It's a complex expression we can't parse
        return {"type": "Other", "value": arg_str}

    def get_tokens(self, body: str) -> List[Tuple[str, str]]:
        """
        Analyzes M-code and returns a list of (token, type) tuples.
        Example: [('let', 'KEYWORD'), ('Source', 'IDENTIFIER'), ...]
        """
        tokens = []
        for mo in TOKEN_REGEX.finditer(body):
            kind = str(mo.lastgroup)
            value = mo.group(kind)
            tokens.append((value, kind))
        return tokens

    def find_dependencies(self, body: str) -> List[str]:
        """
        Parses M-code to find potential dependencies (function calls).
        This is a "best effort" heuristic.
        """
        dependencies = set()

        # 1. Regex to find *any* Library.Function call
        dep_regex = re.compile(
            r'(\b[a-zA-Z_][a-zA-Z0-9_.]*)\s*\(|(#"[^"]+")\s*\(')

        # 2. Regex to *filter out* built-in Library.Function calls
        builtin_func_regex = re.compile(
            r'^[A-Z][a-zA-Z0-9_]*\.[A-Z][a-zA-Z0-9_.]*$')

        for match in dep_regex.finditer(body):
            # name is either the standard call or the quoted call
            name = match.group(1) or match.group(2)
            if name:
                # Exclude keywords
                if name in M_KEYWORDS:
                    continue

                # Exclude data sources
                if name in M_DATA_SOURCES:
                    continue

                if match.group(1) and builtin_func_regex.match(name):
                    continue

                # If it passes all filters, it's a user-defined dependency
                dependencies.add(name)

        return sorted(list(dependencies))

    def find_parameters(self, body: str) -> List[Dict[str, Any]]:
        """
        Finds the parameters of a query if it's a function.
        Returns a list of dicts: [{"name": "...", "type": "...", "optional": ...}]
        """
        params_list = []

        # 1. Regex to find the parameter list (This regex is correct)
        main_match = re.search(
            r'\(\s*(.*?)\s*\)(?:\s+as\s+.*?)?\s*=>',
            body,
            re.DOTALL
        )

        if not main_match:
            return []  # Not a function

        list_str = main_match.group(1)
        if not list_str.strip():
            return []  # Empty function: () => ...

        # 2. STRIP COMMENTS before splitting!
        #    This removes all //... and /*...*/ comments.
        comment_regex = re.compile(r'//[^\n]*|/\*.*?\*/', re.DOTALL)
        cleaned_list_str = re.sub(comment_regex, '', list_str)

        # 3. Split the *cleaned* string by commas
        param_parts = cleaned_list_str.split(',')

        # 4. Regex to parse each part (This regex is correct)
        param_regex = re.compile(
            r'^\s*(optional\s+)?([a-zA-Z0-9_]+)(\s+as\s+(.*))?\s*$',
            re.DOTALL
        )

        for part in param_parts:
            if not part.strip():
                continue

            match = param_regex.match(part)

            if match:
                is_optional = bool(match.group(1))
                name = match.group(2)

                type_str = "any"
                if match.group(4):
                    type_str = match.group(4).strip()

                params_list.append({
                    "name": name,
                    "type": type_str,
                    "optional": is_optional
                })
            else:
                # Fallback (should be hit less often now)
                params_list.append({
                    "name": part.strip(),
                    "type": "unknown",
                    "optional": False
                })

        return params_list

    def find_data_sources(self, body: str) -> List[Dict[str, Any]]:
        """
        Parses M-code to find data source function calls using a token parser
        to correctly handle nested parentheses and newlines.
        """
        sources = []
        tokens = self.get_tokens(body)  # Get all tokens, including whitespace

        i = 0
        while i < len(tokens):
            token_value, token_kind = tokens[i]

            # 1. Is this token a data source function?
            if token_kind == 'DATASOURCE':

                # 2. Look ahead: Is the next non-whitespace token an open-paren?
                j = i + 1
                while j < len(tokens) and tokens[j][1] == 'OTHER' and tokens[j][0].isspace():
                    j += 1  # Skip whitespace

                if j < len(tokens) and tokens[j][0] == '(':
                    # 3. Found a function call! Start capturing the first argument.
                    paren_level = 1
                    arg_start_index = j + 1
                    arg_end_index = -1

                    k = arg_start_index
                    while k < len(tokens):
                        arg_val, arg_kind = tokens[k]

                        if arg_val == '(':
                            paren_level += 1
                        elif arg_val == ')':
                            paren_level -= 1

                        # We found the end of the argument
                        if paren_level == 0 and arg_val == ')':
                            # End on the closing paren (covers single-arg)
                            arg_end_index = k
                            break
                        # We found the end of the *first* argument in a list
                        if paren_level == 1 and arg_val == ',':
                            arg_end_index = k  # End on the comma
                            break

                        k += 1

                    if arg_end_index != -1:
                        # 4. We have the full argument! Reconstruct the string
                        arg_tokens = tokens[arg_start_index:arg_end_index]
                        full_arg_str = "".join(
                            val for val, kind in arg_tokens).strip()

                        # 5. Analyze the argument string
                        ultimate_source = self._find_ultimate_source(
                            full_arg_str)

                        sources.append({
                            "type": token_value,  # e.g., "Csv.Document"
                            # e.g., "File.Contents(filePath)"
                            "full_argument": full_arg_str,
                            # "Variable"
                            "source_type": ultimate_source["type"],
                            # "filePath"
                            "source_value": ultimate_source["value"]
                        })

                        i = k  # Skip ahead to the end of the argument
                        continue  # Restart loop

            i += 1  # Not a data source, check next token

        return sources
