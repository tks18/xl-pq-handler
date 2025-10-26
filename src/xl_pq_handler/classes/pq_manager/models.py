# models.py
import os
from typing import List, Optional, Dict, Any
import yaml
from pydantic import BaseModel, Field, field_validator
from .utils import get_logger

logger = get_logger(__name__)


class PowerQueryMetadata(BaseModel):
    """
    Represents the metadata stored in the index.json
    and in the YAML frontmatter.
    """
    name: str
    category: str = "Uncategorized"
    tags: List[str] = Field(default_factory=list)
    dependencies: List[str] = Field(default_factory=list)
    description: str = ""
    version: str = "1.0"
    path: str  # Absolute path to the .pq file

    @field_validator('name', 'category', 'description', 'version')
    def cleanup_strings(cls, v):
        """Ensure no newlines or weird whitespace in metadata."""
        if v is None:
            return ""
        return str(v).replace("\r", " ").replace("\n", " ").strip()


class PowerQueryScript(BaseModel):
    """
    Represents a full Power Query script,
    including its metadata and M code body.
    """
    meta: PowerQueryMetadata
    body: str

    @classmethod
    def from_file(cls, file_path: str) -> 'PowerQueryScript':
        """Parses a .pq file (with frontmatter) into a model."""
        abs_path = os.path.abspath(file_path)
        logger.debug(f"Parsing file: {abs_path}")

        with open(abs_path, "r", encoding="utf-8") as f:
            text = f.read()

        fm_dict: Dict[str, Any] = {}
        body = text

        if text.lstrip().startswith("---"):
            parts = text.split("---", 2)
            if len(parts) >= 3:
                try:
                    fm_dict = yaml.safe_load(parts[1]) or {}
                    body = parts[2].lstrip("\n")
                except Exception as e:
                    logger.warning(
                        f"Failed to parse YAML frontmatter in {abs_path}: {e}")
                    fm_dict = {}

        # Use file name as fallback for query name
        if "name" not in fm_dict:
            fm_dict["name"] = os.path.splitext(os.path.basename(abs_path))[0]

        fm_dict["path"] = abs_path

        try:
            metadata = PowerQueryMetadata(**fm_dict)
            return cls(meta=metadata, body=body)
        except Exception as e:
            logger.error(f"Failed to validate metadata for {abs_path}: {e}")
            raise

    def to_file_content(self) -> str:
        """Serializes the script back into 'frontmatter + body' format."""
        meta_dict = self.meta.model_dump(exclude={'path'})
        fm = f"---\n{yaml.safe_dump(meta_dict, sort_keys=False, allow_unicode=True)}---\n\n"
        return fm + self.body

    def save(self, overwrite: bool = False) -> None:
        """Saves the script to the path specified in its metadata."""
        target_path = self.meta.path
        if os.path.exists(target_path) and not overwrite:
            raise FileExistsError(
                f"{target_path} already exists. Set overwrite=True."
            )

        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        with open(target_path, "w", encoding="utf-8") as f:
            f.write(self.to_file_content())
        logger.info(f"Saved script to: {target_path}")
