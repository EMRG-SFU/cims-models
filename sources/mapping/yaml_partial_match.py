"""
yaml_partial_match.py

A wrapper class for YAML data that supports partial key matching.
Use in Jupyter notebooks for interactive exploration of YAML data.

Usage:
    from yaml_partial_match import YamlDict, load_yaml
    
    # Load YAML file
    data = load_yaml('merged.yaml')
    
    # Wrap in YamlDict for partial matching
    yaml_data = YamlDict(data)
    
    # Access with partial matches
    sub_tree = yaml_data['emissions']['cumul']  # matches 'emissions_gas' and 'emissions_cumulative'
    
    # Chain access
    result = yaml_data['sect']['Nation']['Oil']
"""

import re
import math
from pathlib import Path
from typing import Any, Dict, List, Union


class YamlDict:
    """
    A dictionary wrapper that supports partial key matching.
    """
    
    def __init__(self, data: Union[Dict, List, Any]):
        self._data = data
        self._is_dict = isinstance(data, dict)
        self._is_list = isinstance(data, list)
    
    @staticmethod
    def _is_blank_or_nan(value: Any) -> bool:
        """Check if a value is None, blank string, or NaN."""
        if value is None:
            return True
        # Check for NaN (works for float('nan') and numpy.nan)
        try:
            if isinstance(value, float) and math.isnan(value):
                return True
        except (TypeError, ValueError):
            pass
        # Check for blank string
        if isinstance(value, str) and not value.strip():
            return True
        return False
    
    def _deep_exact_match(self, data, target):
        """Search the entire tree for a value or key exactly equal to target."""
        if isinstance(data, dict):
            for k, v in data.items():
                # Case 1: key exactly matches
                if str(k) == target:
                    return v
                # Case 2: value exactly matches and is scalar/string
                if isinstance(v, (str, int, float)) and str(v) == target:
                    return v
                # Recurse
                found = self._deep_exact_match(v, target)
                if found is not None:
                    return found
        elif isinstance(data, list):
            for item in data:
                found = self._deep_exact_match(item, target)
                if found is not None:
                    return found
        return None

    def __getitem__(self, key: Union[str, int]) -> 'YamlDict':
        """
        Access items with partial key matching.
        - For dicts: tries exact match first, then bidirectional partial matching
        - For lists: uses integer indices
        - Returns YamlDict wrapper for chaining
        """
        if self._is_list:
            if isinstance(key, int):
                return YamlDict(self._data[key])
            else:
                raise TypeError(f"List indices must be integers, not {type(key).__name__}")
        
        if not self._is_dict:
            raise TypeError(f"'{type(self._data).__name__}' object is not subscriptable")
        
        # Check for blank/empty/NaN key
        if self._is_blank_or_nan(key):
            return None
        
        # Step 1: Try exact match first (case-sensitive)
        if key in self._data:
            return YamlDict(self._data[key])
        
        # Step 2: Try case-insensitive exact match
        key_lower = str(key).lower()
        for k in self._data.keys():
            if str(k).lower() == key_lower:
                return YamlDict(self._data[k])
        
        # Step 3: Deep exact match search anywhere in the YAML tree
        deep = self._deep_exact_match(self._data, key)
        if deep is not None:
            return YamlDict(deep)

        # Step 4: Try bidirectional partial match (case-insensitive substring)
        # Match if key is in dict_key OR dict_key is in key
        matches = []
        for k in self._data.keys():
            k_str = str(k)
            k_lower = k_str.lower()
            # Bidirectional: key in dict_key OR dict_key in key
            if key_lower in k_lower or k_lower in key_lower:
                matches.append(k)
        
        if len(matches) == 0:
            # Show available keys to help debug
            available = list(self._data.keys())[:5]  # Show first 5 keys
            available_str = ", ".join(f"'{k}'" for k in available)
            if len(self._data) > 5:
                available_str += ", ..."
            raise KeyError(f"No key found matching '{key}'. Available keys: {available_str}")
        elif len(matches) == 1:
            return YamlDict(self._data[matches[0]])
        else:
            # Multiple partial matches - return dict with all matches
            result = {k: self._data[k] for k in matches}
            return YamlDict(result)
    
    def __repr__(self) -> str:
        """String representation."""
        return f"YamlDict({self._data!r})"
    
    def __str__(self) -> str:
        """Pretty print the data."""
        return self._pretty_print(self._data, indent=0)
    
    def _pretty_print(self, data: Any, indent: int = 0) -> str:
        """Pretty print YAML-like structure."""
        pad = "  " * indent
        
        if isinstance(data, dict):
            if not data:
                return "{}"
            lines = []
            for k, v in data.items():
                if isinstance(v, (dict, list)) and v:
                    lines.append(f"{pad}{k}:")
                    lines.append(self._pretty_print(v, indent + 1))
                else:
                    lines.append(f"{pad}{k}: {v}")
            return "\n".join(lines)
        elif isinstance(data, list):
            if not data:
                return "[]"
            lines = []
            for item in data:
                if isinstance(item, (dict, list)):
                    lines.append(f"{pad}-")
                    lines.append(self._pretty_print(item, indent + 1))
                else:
                    lines.append(f"{pad}- {item}")
            return "\n".join(lines)
        else:
            return f"{pad}{data}"
    
    def get_data(self) -> Any:
        """Return the underlying data."""
        return self._data
    
    def keys(self) -> List[str]:
        """Return all keys if this is a dict."""
        if self._is_dict:
            return list(self._data.keys())
        return []
    
    def search(self, term: str) -> 'YamlDict':
        """
        Search for keys containing the term anywhere in the tree.
        Returns a new YamlDict with matching results.
        
        Returns None if blank/NaN search term provided.
        """
        # Check for blank/empty/NaN term
        if self._is_blank_or_nan(term):
            return None
        
        results = self._search_recursive(self._data, str(term).lower())
        if not results:
            raise KeyError(f"No keys found matching '{term}'")
        return YamlDict(results)
    
    def _search_recursive(self, data: Any, term: str) -> Dict:
        """Recursively search for matching keys."""
        results = {}
        
        if isinstance(data, dict):
            for key, value in data.items():
                key_str = str(key)
                if term in key_str.lower():
                    results[key] = value
                elif isinstance(value, dict):
                    child_results = self._search_recursive(value, term)
                    if child_results:
                        results[key] = child_results
        
        return results
    
    def to_dict(self) -> Dict:
        """Convert to a regular dictionary (recursive)."""
        if self._is_dict:
            return {k: YamlDict(v).to_dict() if isinstance(v, (dict, list)) else v 
                    for k, v in self._data.items()}
        elif self._is_list:
            return [YamlDict(item).to_dict() if isinstance(item, (dict, list)) else item 
                    for item in self._data]
        else:
            return self._data
    
    def to_number(self, default: float = 0.0) -> float:
        """
        Convert the data to a number. Useful for extracting numeric values.
        
        Args:
            default: Default value to return if conversion fails
            
        Returns:
            Float value or default
        """
        try:
            if isinstance(self._data, (int, float)):
                return float(self._data)
            if isinstance(self._data, str):
                # Try to parse as number
                cleaned = self._data.strip()
                # Remove commas if present
                cleaned = cleaned.replace(',', '')
                return float(cleaned)
            return default
        except (ValueError, TypeError, AttributeError):
            return default
    
    def find_key(self, target_key: str) -> 'YamlDict':
        """
        Recursively search for a key anywhere in the structure.
        Uses partial matching.
        
        Args:
            target_key: The key name to find (supports partial matching)
        
        Returns:
            YamlDict containing the data at that key, or None if blank/NaN key provided
            
        Raises:
            KeyError if not found
        """
        # Check for blank/empty/NaN key
        if self._is_blank_or_nan(target_key):
            return None
        
        result = self._find_key_recursive(self._data, target_key)
        if result is None:
            raise KeyError(f"Key '{target_key}' not found in data structure")
        return YamlDict(result)
    
    def _find_key_recursive(self, data: Any, target_key: str) -> Any:
        """Helper for recursive key search with bidirectional partial matching."""
        if not isinstance(data, dict):
            return None
        
        target_lower = target_key.lower()
        
        # Try exact match first (case-sensitive)
        if target_key in data:
            return data[target_key]
        
        # Try case-insensitive exact match
        for k in data.keys():
            if str(k).lower() == target_lower:
                return data[k]
        
        # Try bidirectional partial match at this level only
        matches = []
        for k in data.keys():
            k_lower = str(k).lower()
            # Bidirectional: target in key OR key in target
            if target_lower in k_lower or k_lower in target_lower:
                matches.append(k)
        
        # If there's a partial match, do NOT return it yet.
        # First check recursively for deeper exact matches.
        if len(matches) >= 1:
            # Search all children for deeper exact match
            for value in data.values():
                result = self._find_key_recursive(value, target_key)
                if result is not None:
                    return result

            # If no deeper exact match exists, THEN:
            if len(matches) == 1:
                return data[matches[0]]
            else:
                # Multiple partial matches: return all
                return {k: data[k] for k in matches}
        
        # Recursively search in all values
        for value in data.values():
            result = self._find_key_recursive(value, target_key)
            if result is not None:
                return result
        
        # If we got here and had multiple matches at this level, return them
        if len(matches) > 1:
            return {k: data[k] for k in matches}
        
        return None
    
    def find_leaf_keys(self, parent_key: str = "") -> List[Dict]:
        """
        Find all keys that have "include" fields.
        Stops searching a branch once found.
        Automatically searches for matching units in the "unit" branch and applies multiplier.
        
        Args:
            parent_key: The name of the parent key (for tracking)
        
        Returns:
            List of dictionaries containing key, include value, and unit info with multiplied value
        """
        return self._find_leaf_keys_recursive(self._data, "", parent_key)
    
    def _find_leaf_keys_recursive(self, data: Any, current_path: str, parent_key: str) -> List[Dict]:
        """Helper for finding leaf keys with include and automatic unit lookup."""
        results = []
        
        if not isinstance(data, dict):
            return results
        
        has_operation = False
        
        # Check if this dict has an "include" key
        if "include" in data:
            temp = {}
            for k, v in data.items():
                temp[k] = v
            temp["key"] = parent_key
            
            # Search for matching unit in the "unit" branch of the root data
            unit_multiplier = 1.0  # Default multiplier
            if parent_key:
                try:
                    # Try to find "unit" branch at root level
                    root_data = self._get_root_data()
                    if isinstance(root_data, dict) and "unit" in root_data:
                        unit_branch = root_data["unit"]
                        # Search for matching parent key in unit branch
                        unit_data = self._find_unit_recursive(unit_branch, parent_key)
                        if unit_data and isinstance(unit_data, dict):
                            # Add unit information to the result
                            for uk, uv in unit_data.items():
                                temp[f"unit_{uk}"] = uv
                            # Get multiplier if present
                            if "multiplier" in unit_data:
                                try:
                                    unit_multiplier = float(unit_data["multiplier"])
                                except (ValueError, TypeError):
                                    unit_multiplier = 1.0
                except (KeyError, AttributeError):
                    # No unit branch found, use default
                    pass
            
            # Store the multiplier
            temp["unit_multiplier"] = unit_multiplier
            
            results.append(temp)
            has_operation = True
        
        # If we found include, stop searching this branch
        if has_operation:
            return results
        
        # Recursively search through all keys
        for key, value in data.items():
            if key in ["include"]:
                continue
            
            new_path = f"{current_path}.{key}" if current_path else key
            results.extend(self._find_leaf_keys_recursive(value, new_path, key))
        
        return results
    
    def _get_root_data(self) -> Any:
        """Get the root data structure. Override this if needed."""
        return self._data
    
    def _find_unit_recursive(self, unit_data: Any, target_key: str) -> Any:
        """Search for matching key in unit branch using partial matching."""
        if not isinstance(unit_data, dict):
            return None
        
        target_lower = target_key.lower()
        
        # Try exact match first
        if target_key in unit_data:
            return unit_data[target_key]
        
        # Try case-insensitive exact match
        for k in unit_data.keys():
            if str(k).lower() == target_lower:
                return unit_data[k]
        
        # Try bidirectional partial match
        matches = []
        for k in unit_data.keys():
            k_lower = str(k).lower()
            if target_lower in k_lower or k_lower in target_lower:
                matches.append(k)
        
        if len(matches) == 1:
            return unit_data[matches[0]]
        elif len(matches) > 1:
            # Multiple matches - return first one
            return unit_data[matches[0]]
        
        # Recursively search in nested structures
        for value in unit_data.values():
            if isinstance(value, dict):
                result = self._find_unit_recursive(value, target_key)
                if result is not None:
                    return result
        
        return None
    
    def find_from_key(self, start_key: str) -> List[Dict]:
        """
        Find all child keys with "include" field starting from a given key.
        Uses partial matching to find the start key.
        Automatically searches for matching units in the "unit" branch and applies multiplier.
        
        Args:
            start_key: The key name to start searching from (supports partial matching)
        
        Returns:
            List of dictionaries containing key name, include value, and unit info with multiplier
            Returns empty list if blank/NaN key provided
        """
        # Check for blank/empty/NaN key
        if self._is_blank_or_nan(start_key):
            return []
        
        try:
            start_data = self.find_key(start_key)
            if start_data is None:
                return []
            return start_data.find_leaf_keys(start_key)
        except KeyError:
            print(f"Key '{start_key}' not found in data")
            return []


# ----------------------------------------------------------------------
# YAML Parsing Functions
# ----------------------------------------------------------------------

def _indent_of(line: str) -> int:
    return len(line) - len(line.lstrip(" "))

def _is_blank(line: str) -> bool:
    return not line.strip()

def _is_comment(line: str) -> bool:
    return line.lstrip().startswith("#")

def _parse_scalar(text: str) -> Any:
    txt = text.strip()
    if txt in {"null", "~"}:
        return None
    if txt.lower() == "true":
        return True
    if txt.lower() == "false":
        return False
    if re.fullmatch(r"-?\d+", txt):
        return int(txt)
    if re.fullmatch(r"-?\d+(?:\.\d+)?", txt):
        return float(txt)
    if (txt.startswith('"') and txt.endswith('"')) or (
        txt.startswith("'") and txt.endswith("'")
    ):
        return txt[1:-1]
    return txt

def _parse_mapping(lines: List[str], start: int, parent_indent: int) -> tuple:
    mapping = {}
    i = start

    while i < len(lines):
        line = lines[i]

        if _is_blank(line) or _is_comment(line):
            i += 1
            continue

        cur_indent = _indent_of(line)
        if cur_indent < parent_indent:
            break

        kv_match = re.match(r"\s*([^:#]+?)\s*:\s*(.*)$", line)
        if not kv_match:
            i += 1
            continue

        raw_key, raw_val = kv_match.groups()
        key = raw_key.strip()

        if "#" in raw_val:
            raw_val = raw_val.split("#")[0]
        raw_val = raw_val.strip()

        if raw_val == "":
            next_i = i + 1
            if next_i >= len(lines):
                mapping[key] = {}
                i += 1
                continue
            next_line = lines[next_i].lstrip()
            sub_indent = _indent_of(lines[next_i])

            if next_line.startswith("-"):
                lst, new_i = _parse_list(lines, next_i, sub_indent)
                mapping[key] = lst
                i = new_i
            else:
                sub_map, new_i = _parse_mapping(lines, next_i, sub_indent)
                mapping[key] = sub_map
                i = new_i
            continue

        if raw_val.startswith("[") and raw_val.endswith("]"):
            list_content = raw_val[1:-1]
            items = [_parse_scalar(item.strip()) for item in list_content.split(",") if item.strip()]
            mapping[key] = items
            i += 1
            continue

        mapping[key] = _parse_scalar(raw_val)
        i += 1

    return mapping, i

def _parse_list(lines: List[str], start: int, parent_indent: int) -> tuple:
    result = []
    i = start

    while i < len(lines):
        line = lines[i]

        if _is_blank(line) or _is_comment(line):
            i += 1
            continue

        cur_indent = _indent_of(line)
        if cur_indent < parent_indent:
            break

        stripped = line.lstrip()
        if not stripped.startswith("-"):
            break

        item_val = stripped[1:].strip()

        if "#" in item_val:
            item_val = item_val.split("#")[0].strip()

        if item_val == "":
            next_i = i + 1
            if next_i >= len(lines):
                result.append({})
                i += 1
                continue

            next_line = lines[next_i].lstrip()
            sub_indent = _indent_of(lines[next_i])

            if next_line.startswith("-"):
                nested_list, new_i = _parse_list(lines, next_i, sub_indent)
                result.append(nested_list)
                i = new_i
            else:
                nested_map, new_i = _parse_mapping(lines, next_i, sub_indent)
                result.append(nested_map)
                i = new_i
            continue

        result.append(_parse_scalar(item_val))
        i += 1

    return result, i

def load_yaml(filepath: Union[str, Path]) -> YamlDict:
    """
    Load a YAML file and return a YamlDict for partial key matching.
    
    Args:
        filepath: Path to the YAML file
        
    Returns:
        YamlDict object with partial matching support
    """
    filepath = Path(filepath)
    raw_lines = filepath.read_text(encoding="utf-8").splitlines()
    data, _ = _parse_mapping(raw_lines, 0, 0)
    return YamlDict(data)


# Example usage for Jupyter notebooks
if __name__ == "__main__":
    print("This module is meant to be imported in a Jupyter notebook.")
    print("\nExample usage:")
    print("  from yaml_partial_match import load_yaml")
    print("  data = load_yaml('merged.yaml')")
    print("  sub_tree = data['sect']['Nation']['Oil']")
    print("  print(sub_tree)")