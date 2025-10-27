# dependencies.py
from typing import List, Dict, Set, Optional, Any
from .models import PowerQueryMetadata
from .storage import PQFileStore
from .utils import get_logger

logger = get_logger(__name__)


class CircularDependencyError(Exception):
    """Raised when a circular dependency is detected."""
    pass


class DependencyResolver:
    """
    Resolves dependency graphs for Power Query scripts.
    Provides correct insertion order and detects cycles.
    """

    def __init__(self, store: PQFileStore):
        self.store = store

    def get_insertion_order(self, names: List[str]) -> List[str]:
        """
        Returns a list of all required query names (including dependencies)
        in the correct order for insertion.

        Uses topological sort.
        """
        resolved: Set[str] = set()  # Queries in the final sorted list
        visiting: Set[str] = set()  # Queries currently in the recursion stack
        graph: Dict[str, List[str]] = {}  # Adjacency list

        # 1. Build the full dependency graph for all items in the store
        # This is more robust as it can find dependencies of dependencies
        self.store.load_index()  # Ensure index is fresh
        all_metadata = self.store._index.values()

        for meta in all_metadata:
            graph[meta.name] = meta.dependencies

        # 2. Define the recursive Depth First Search (DFS) function
        def dfs(name: str):
            if name in resolved:
                return  # Already processed

            if name in visiting:
                raise CircularDependencyError(
                    f"Circular dependency detected involving: {name}")

            meta = self.store.get_metadata_by_name(name)
            if not meta:
                # This is a missing dependency
                logger.warning(
                    f"Missing dependency: '{name}' not found in store.")
                # We can't proceed with this branch, but we don't error out
                # The caller (e.g., insertion) will fail if it's truly needed
                return

            visiting.add(name)
            for dep_name in meta.dependencies:
                dfs(dep_name)  # Recurse

            visiting.remove(name)
            resolved.add(name)

        # 3. Run DFS on all initially requested names
        for name in names:
            if name not in resolved:
                dfs(name)

        # 4. Filter and sort the final list
        # 'resolved' now contains *all* dependencies
        # We need to return them in an order that respects the graph
        # The 'resolved' set isn't ordered, so we must re-build

        # A full topological sort (Kahn's algorithm or post-order DFS)
        # is needed for the *correct* order.

        # Let's re-do this with a proper post-order DFS to get the order

        ordered_list: List[str] = []
        visited: Set[str] = set()

        def dfs_for_order(name: str):
            visited.add(name)

            meta = self.store.get_metadata_by_name(name)
            if not meta:
                logger.warning(f"Skipping missing dependency: {name}")
                return

            for dep_name in meta.dependencies:
                if dep_name not in visited:
                    dfs_for_order(dep_name)

            ordered_list.append(name)  # Add *after* dependencies are visited

        # Check for cycles first (using the first DFS)
        try:
            for name in names:
                if name not in resolved:  # 'resolved' is from the cycle check
                    dfs(name)
        except CircularDependencyError as e:
            logger.error(f"Failed to resolve dependencies: {e}")
            raise  # Re-raise the error

        # If no cycles, run the ordering DFS
        for name in names:
            if name not in visited:
                dfs_for_order(name)

        # `ordered_list` now contains the requested names *and* their dependencies,
        # with dependencies appearing *before* the queries that need them.

        # We just need to filter this list to only include the items
        # from the original request + their dependencies.
        # The 'resolved' set from the cycle check has exactly this.

        final_ordered_list = [
            name for name in ordered_list if name in resolved]

        return final_ordered_list

    def get_dependency_tree(self, name: str) -> Dict[str, Any]:
        """
        Builds a recursive dictionary representing the dependency
        tree for a single query.
        Returns a dict like: {"name": "A", "children": [{"name": "B"}]}
        """
        tree = {"name": name, "children": []}

        # Use a set to avoid infinite loops on circular dependencies
        visited = set()

        def build_tree(current_node_dict, current_name):
            if current_name in visited:
                return  # Stop recursion

            visited.add(current_name)

            meta = self.store.get_metadata_by_name(current_name)
            if not meta or not meta.dependencies:
                return  # No more children

            for dep_name in meta.dependencies:
                child_node_dict = {"name": dep_name, "children": []}
                current_node_dict["children"].append(child_node_dict)
                # Recurse
                build_tree(child_node_dict, dep_name)

        build_tree(tree, name)
        return tree
