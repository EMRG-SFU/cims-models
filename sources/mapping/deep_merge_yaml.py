"""
deep_merge_yaml.py

Read multiple YAML-like files that use anchors (&), aliases (*), and the merge key (<<:).
Perform a *deep* merge of the files and their anchored mappings without using any external YAML library.

Usage:
    python deep_merge_yaml.py input1.yaml [input2.yaml ...] [-o output.yaml] [--add-multipliers]
    
    If -o is not specified, outputs to stdout.
    --add-multipliers: Add multiplier keys recursively to all nodes with 'include' keys
"""

import sys
import re
import copy
from pathlib import Path
from typing import List, Tuple, Dict, Any


# ----------------------------------------------------------------------
# Helper: pretty-print a Python value as a YAML scalar (very small subset)
# ----------------------------------------------------------------------
def yaml_scalar(val: Any) -> str:
    """Return a YAML-compatible representation of a scalar."""
    if val is None:
        return "null"
    if isinstance(val, bool):
        return "true" if val else "false"
    if isinstance(val, (int, float)):
        return str(val)
    escaped = str(val).replace('"', r'\"')
    return f'"{escaped}"'


# ----------------------------------------------------------------------
# Deep-merge implementation
# ----------------------------------------------------------------------
def has_unit_key(obj: Any) -> bool:
    """
    Recursively check if a dictionary (or nested dicts) contains a 'Unit' key.
    """
    if not isinstance(obj, dict):
        return False
    if 'Unit' in obj:
        return True
    return any(has_unit_key(v) for v in obj.values())


def extract_unit_structure(obj: Any, path: List[str] = None) -> Dict[Any, Any]:
    """
    Extract only the branches that lead to 'Unit' keys.
    Returns a dictionary containing only the paths to Unit keys.
    """
    if path is None:
        path = []
    
    if not isinstance(obj, dict):
        return {}
    
    result = {}
    
    for key, value in obj.items():
        if key == 'Unit':
            # Found a Unit key, include it
            result[key] = copy.deepcopy(value)
        elif isinstance(value, dict) and has_unit_key(value):
            # This branch leads to a Unit, recursively extract it
            sub_unit = extract_unit_structure(value, path + [key])
            if sub_unit:
                result[key] = sub_unit
    
    return result


def deep_merge(base: Dict[Any, Any], overlay: Dict[Any, Any]) -> Dict[Any, Any]:
    """
    Recursively merge ``overlay`` into ``base`` and return ``base``.
    * Scalars in ``overlay`` replace entries in ``base``.
    * ``None`` deletes the key from ``base``, BUT preserves Unit dictionaries.
    * Nested dicts are merged recursively.
    """
    for key, ov_val in overlay.items():
        if ov_val is None:
            # Check if the base value has a Unit key somewhere in its tree
            base_val = base.get(key)
            if base_val and has_unit_key(base_val):
                # Preserve only the Unit structure
                unit_structure = extract_unit_structure(base_val)
                base[key] = unit_structure
            else:
                # No Unit key, safe to delete
                base.pop(key, None)
            continue

        if isinstance(ov_val, dict):
            base_val = base.get(key)
            if not isinstance(base_val, dict):
                base_val = {}
                base[key] = base_val
            deep_merge(base_val, ov_val)
        else:
            base[key] = ov_val
    return base


def apply_replacements(data: Any, replacements: Dict[str, Any]) -> None:
    """
    Recursively search through data and replace matching keys with new values IN-PLACE.
    If the key doesn't exist but the node has an 'include' key, ADD it.
    
    Args:
        data: The data structure to search and modify
        replacements: Dictionary of {key_to_find: new_value}
    """
    if isinstance(data, dict):
        # First, handle existing keys that need replacement
        for key, value in list(data.items()):
            # Check if this key should be replaced
            if key in replacements:
                # Replace the entire value
                data[key] = replacements[key]
                print(f"DEBUG: Replaced {key} = {replacements[key]}")
            elif isinstance(value, dict):
                # Recursively apply replacements to nested dicts
                apply_replacements(value, replacements)
            elif isinstance(value, list):
                # Recursively apply replacements to lists
                apply_replacements(value, replacements)
        
        # Second, if this node has 'include' and is missing replacement keys, add them
        if 'include' in data:
            for replace_key, replace_value in replacements.items():
                if replace_key not in data:
                    data[replace_key] = replace_value
                    print(f"DEBUG: Added {replace_key} = {replace_value} (was missing)")
                    
    elif isinstance(data, list):
        for item in data:
            if isinstance(item, (dict, list)):
                apply_replacements(item, replacements)


def add_multipliers_recursive(data: Any, parent_multiplier: float = 1.0, is_root: bool = True) -> Any:
    """
    Recursively add multiplier keys to all nodes that have an 'include' key.
    Removes multiplier keys from nodes without 'include' keys.
    Inherits multiplier from parent unless a new multiplier is specified.
    Resets to default multiplier (1.0) at each top-level branch.
    DOES NOT overwrite existing multipliers - only adds where missing.
    
    Args:
        data: The data structure (dict, list, or scalar)
        parent_multiplier: The multiplier value from the parent node
        is_root: True if this is a root-level node
    
    Returns:
        Modified data with multipliers added only to nodes with 'include' keys
    """
    if not isinstance(data, dict):
        return data
    
    # At root level, process each top-level key independently
    if is_root:
        for key, value in list(data.items()):
            if isinstance(value, dict):
                # Reset multiplier to 1.0 for each top-level branch
                add_multipliers_recursive(value, parent_multiplier=1.0, is_root=False)
        return data
    
    # Determine current multiplier for this level
    if 'multiplier' in data:
        current_multiplier = data['multiplier']
        # Remove multiplier from this node if it doesn't have 'include'
        if 'include' not in data:
            # But keep it temporarily to pass to children
            # Don't delete it yet
            pass
    else:
        current_multiplier = parent_multiplier
    
    # Add multiplier only if this node has an 'include' key AND doesn't already have one
    if 'include' in data and 'multiplier' not in data:
        data['multiplier'] = current_multiplier
    
    # Process all children with the current multiplier
    for key, value in list(data.items()):
        if isinstance(value, dict):
            # Recursively process children (not at root level)
            add_multipliers_recursive(value, current_multiplier, is_root=False)
        elif isinstance(value, list):
            # Process list items
            for item in value:
                if isinstance(item, dict):
                    add_multipliers_recursive(item, current_multiplier, is_root=False)
    
    # NOW remove multiplier from this node if it doesn't have 'include'
    # Do this after processing children so they inherit it first
    if 'multiplier' in data and 'include' not in data:
        del data['multiplier']
    
    return data


# ----------------------------------------------------------------------
# Parsing utilities
# ----------------------------------------------------------------------
def _indent_of(line: str) -> int:
    """Number of leading spaces (tabs are not supported)."""
    return len(line) - len(line.lstrip(" "))


def _is_blank(line: str) -> bool:
    return not line.strip()


def _is_comment(line: str) -> bool:
    """Check if a line is a comment (starts with #)."""
    return line.lstrip().startswith("#")


def _parse_scalar(text: str) -> Any:
    """Very small scalar parser – handles null, booleans, numbers, strings."""
    txt = text.strip()
    if txt in {"null", "~"}:
        return None
    if txt.lower() == "true":
        return True
    if txt.lower() == "false":
        return False
    # Integer?
    if re.fullmatch(r"-?\d+", txt):
        return int(txt)
    # Float (including scientific notation)?
    if re.fullmatch(r"-?\d+\.?\d*[eE][+-]?\d+", txt):
        return float(txt)
    if re.fullmatch(r"-?\d+\.\d+", txt):
        return float(txt)
    # Strip quotes if present
    if (txt.startswith('"') and txt.endswith('"')) or (
        txt.startswith("'") and txt.endswith("'")
    ):
        return txt[1:-1]
    return txt


def parse_mapping(
    lines: List[str], start: int, parent_indent: int, anchors: Dict[str, Any]
) -> Tuple[Dict[Any, Any], int]:
    """
    Parse a block of lines that represents a mapping.
    Returns (mapping, index_of_next_unparsed_line).
    """
    mapping: Dict[Any, Any] = {}
    merge_replace_directives: Dict[str, Any] = {}  # Collect merge_replace directives
    i = start

    while i < len(lines):
        line = lines[i]

        if _is_blank(line) or _is_comment(line):
            i += 1
            continue

        cur_indent = _indent_of(line)
        if cur_indent < parent_indent:
            break

        # --------------------------------------------------------------
        # 1️⃣  Detect a merge key:   <<: *anchor   or  <<: [*a, *b]
        #     with optional merge_include, merge_exclude, or merge_replace modifiers
        # --------------------------------------------------------------
        m_merge = re.match(r"\s*<<\s*:\s*(.+)$", line)
        if m_merge:
            merge_val = m_merge.group(1).strip()
            
            # Check for merge_include, merge_exclude, or merge_replace modifiers
            merge_include = None
            merge_exclude = None
            merge_replace = {}  # Dictionary of key: value pairs to replace
            
            # Look ahead for merge directives
            next_i = i + 1
            while next_i < len(lines):
                next_line = lines[next_i]
                if _is_blank(next_line) or _is_comment(next_line):
                    next_i += 1
                    continue
                
                next_indent = _indent_of(next_line)
                # Directives must be at the same indent level as the << line
                if next_indent < cur_indent:
                    break
                if next_indent > cur_indent:
                    # This is content at a deeper level, stop looking for directives
                    break
                
                # Check for merge_include directive
                m_include = re.match(r"\s*merge_include\s*:\s*(.+)$", next_line)
                if m_include:
                    merge_include = m_include.group(1).strip().strip('"\'')
                    next_i += 1
                    continue
                
                # Check for merge_exclude directive
                m_exclude = re.match(r"\s*merge_exclude\s*:\s*(.+)$", next_line)
                if m_exclude:
                    merge_exclude = m_exclude.group(1).strip().strip('"\'')
                    next_i += 1
                    continue
                
                # Check for merge_replace directive (can have multiple)
                m_replace = re.match(r"\s*merge_replace\s*:\s*(.+)$", next_line)
                if m_replace:
                    replace_spec = m_replace.group(1).strip()
                    # Strip comments from replace_spec
                    if "#" in replace_spec:
                        replace_spec = replace_spec.split("#")[0].strip()
                    # Parse key: value format
                    replace_match = re.match(r"([^:]+):\s*(.+)$", replace_spec)
                    if replace_match:
                        replace_key = replace_match.group(1).strip().strip('"\'')
                        replace_val = _parse_scalar(replace_match.group(2).strip())
                        merge_replace[replace_key] = replace_val
                    next_i += 1
                    continue
                
                # If we hit another key at the same level that's not a directive, stop
                break
            
            # Update i to skip processed directives
            if next_i > i + 1:
                i = next_i
            else:
                i += 1
            
            # Process anchor references
            anchor_names = re.findall(r"\*(\w+)", merge_val)
            if not anchor_names:
                raise ValueError(f"Invalid merge syntax on line {i+1}: {line.strip()}")
            
            # Deep merge each anchor into the mapping
            for anchor_name in anchor_names:
                if anchor_name not in anchors:
                    raise ValueError(f"Undefined anchor '{anchor_name}' used for merge.")
                
                anchor_data = copy.deepcopy(anchors[anchor_name])
                
                # Apply merge_include filter
                if merge_include:
                    if merge_include in anchor_data:
                        anchor_data = {merge_include: anchor_data[merge_include]}
                    else:
                        anchor_data = {}
                
                # Apply merge_exclude filter
                if merge_exclude:
                    anchor_data.pop(merge_exclude, None)
                
                # Use deep_merge (don't apply replacements yet)
                deep_merge(mapping, anchor_data)
            
            # Apply merge_replace filters AFTER merging into mapping
            if merge_replace:
                apply_replacements(mapping, merge_replace)
            
            continue

        # --------------------------------------------------------------
        # 2️⃣  Normal key/value line
        # --------------------------------------------------------------
        kv_match = re.match(r"\s*([^:#]+?)\s*:\s*(.*)$", line)
        if not kv_match:
            raise ValueError(f"Cannot parse line {i+1}: {line!r}")

        raw_key, raw_val = kv_match.groups()
        key = raw_key.strip()
        
        # Strip comments (everything after #)
        if "#" in raw_val:
            raw_val = raw_val.split("#")[0]
        raw_val = raw_val.strip()
        
        # Check for scoped merge_replace directive
        if key == "merge_replace":
            replace_spec = raw_val
            # Parse key: value format
            replace_match = re.match(r"([^:]+):\s*(.+)$", replace_spec)
            if replace_match:
                replace_key = replace_match.group(1).strip().strip('"\'')
                replace_val = _parse_scalar(replace_match.group(2).strip())
                # Collect the replacement to apply at the end
                merge_replace_directives[replace_key] = replace_val
                print(f"DEBUG: Collected merge_replace: {replace_key} = {replace_val} at indent {cur_indent}")
            i += 1
            continue

        # --------------------------------------------------------------
        # 3️⃣  Handle an anchor attached to the key:   key: &anchor
        # --------------------------------------------------------------
        m_anchor = re.match(r"&(\w+)\s*$", raw_val)
        if m_anchor:
            anchor_name = m_anchor.group(1)
            # Find indent level of the next line
            next_i = i + 1
            if next_i >= len(lines):
                raise ValueError(f"Unexpected end of file after anchor {anchor_name}")
            next_indent = _indent_of(lines[next_i])
            sub_map, new_i = parse_mapping(lines, next_i, next_indent, anchors)
            anchors[anchor_name] = sub_map
            mapping[key] = sub_map
            i = new_i
            continue

        # --------------------------------------------------------------
        # 4️⃣  Empty value → could be nested mapping or list
        # --------------------------------------------------------------
        if raw_val == "":
            next_i = i + 1
            if next_i >= len(lines):
                mapping[key] = {}
                i += 1
                continue
            next_line = lines[next_i].lstrip()
            sub_indent = _indent_of(lines[next_i])
            
            # Check if next line is a list item
            if next_line.startswith("-"):
                lst, new_i = parse_list(lines, next_i, sub_indent, anchors)
                mapping[key] = lst
                i = new_i
            else:
                sub_map, new_i = parse_mapping(lines, next_i, sub_indent, anchors)
                
                # If key already exists and both are dicts, deep merge instead of replace
                if key in mapping and isinstance(mapping[key], dict) and isinstance(sub_map, dict):
                    deep_merge(mapping[key], sub_map)
                else:
                    mapping[key] = sub_map
                i = new_i
            continue

        # --------------------------------------------------------------
        # 5️⃣  Alias as a value: key: *anchor
        # --------------------------------------------------------------
        m_alias = re.match(r"\*(\w+)$", raw_val)
        if m_alias:
            alias_name = m_alias.group(1)
            if alias_name not in anchors:
                raise ValueError(f"Undefined alias '{alias_name}' on line {i+1}")
            mapping[key] = copy.deepcopy(anchors[alias_name])
            i += 1
            continue

        # --------------------------------------------------------------
        # 6️⃣  Inline list: key: [item1, item2]
        # --------------------------------------------------------------
        if raw_val.startswith("[") and raw_val.endswith("]"):
            list_content = raw_val[1:-1]
            items = [_parse_scalar(item.strip()) for item in list_content.split(",") if item.strip()]
            mapping[key] = items
            i += 1
            continue

        # --------------------------------------------------------------
        # 7️⃣  Plain scalar value
        # --------------------------------------------------------------
        mapping[key] = _parse_scalar(raw_val)
        i += 1

    # Apply any collected merge_replace directives to this mapping scope
    if merge_replace_directives:
        print(f"DEBUG: Applying merge_replace at indent {parent_indent}: {merge_replace_directives}")
        print(f"DEBUG: Mapping structure:")
        import json
        print(json.dumps(mapping, indent=2, default=str)[:500])  # First 500 chars
        apply_replacements(mapping, merge_replace_directives)

    return mapping, i


def parse_list(
    lines: List[str], start: int, parent_indent: int, anchors: Dict[str, Any]
) -> Tuple[List[Any], int]:
    """
    Parse a block of lines that represents a list.
    Returns (list, index_of_next_unparsed_line).
    """
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

        # Get the value after the dash
        item_val = stripped[1:].strip()

        # Strip comments
        if "#" in item_val:
            item_val = item_val.split("#")[0].strip()

        # Empty value after dash → nested structure
        if item_val == "":
            next_i = i + 1
            if next_i >= len(lines):
                result.append({})
                i += 1
                continue
            
            next_line = lines[next_i].lstrip()
            sub_indent = _indent_of(lines[next_i])
            
            if next_line.startswith("-"):
                # Nested list
                nested_list, new_i = parse_list(lines, next_i, sub_indent, anchors)
                result.append(nested_list)
                i = new_i
            else:
                # Nested mapping
                nested_map, new_i = parse_mapping(lines, next_i, sub_indent, anchors)
                result.append(nested_map)
                i = new_i
            continue

        # Alias reference
        m_alias = re.match(r"\*(\w+)$", item_val)
        if m_alias:
            alias_name = m_alias.group(1)
            if alias_name not in anchors:
                raise ValueError(f"Undefined alias '{alias_name}' on line {i+1}")
            result.append(copy.deepcopy(anchors[alias_name]))
            i += 1
            continue

        # Scalar value
        result.append(_parse_scalar(item_val))
        i += 1

    return result, i


# ----------------------------------------------------------------------
# YAML dumper (tiny subset)
# ----------------------------------------------------------------------
def dump_yaml(data: Any, indent: int = 0) -> str:
    pad = " " * indent
    if isinstance(data, dict):
        lines = []
        for k, v in data.items():
            if isinstance(v, dict) and v:  # Non-empty dict
                lines.append(f"{pad}{k}:")
                lines.append(dump_yaml(v, indent + 4))
            elif isinstance(v, dict) and not v:  # Empty dict
                lines.append(f"{pad}{k}:")
            elif isinstance(v, list):  # List
                lines.append(f"{pad}{k}:")
                lines.append(dump_yaml(v, indent + 4))
            else:  # Scalar value
                lines.append(f"{pad}{k}: {yaml_scalar(v)}")
        return "\n".join(lines)
    elif isinstance(data, list):
        lines = []
        for item in data:
            if isinstance(item, (dict, list)):
                lines.append(f"{pad}-")
                lines.append(dump_yaml(item, indent + 4))
            else:
                lines.append(f"{pad}- {yaml_scalar(item)}")
        return "\n".join(lines)
    else:
        return f"{pad}{yaml_scalar(data)}"


# ----------------------------------------------------------------------
# Main driver
# ----------------------------------------------------------------------
def parse_file(filepath: Path, anchors: Dict[str, Any]) -> Dict[Any, Any]:
    """Parse a single YAML file and return the document."""
    raw_lines = filepath.read_text(encoding="utf-8").splitlines()
    doc, _ = parse_mapping(raw_lines, 0, 0, anchors)
    return doc


def main(filepaths: List[Path], add_multipliers: bool = False) -> str:
    """Parse and merge multiple YAML files, keeping only top-level keys from the last file."""
    anchors: Dict[str, Any] = {}
    merged_doc: Dict[Any, Any] = {}
    
    # Parse all files to build up anchors and merge data
    all_docs = []
    for filepath in filepaths:
        doc = parse_file(filepath, anchors)
        all_docs.append(doc)
        deep_merge(merged_doc, doc)
    
    # If multiple files, keep only the top-level keys from the last file
    if len(all_docs) > 1:
        last_doc = all_docs[-1]
        final_doc = {}
        for key in last_doc.keys():
            if key in merged_doc:
                final_doc[key] = merged_doc[key]
        merged_doc = final_doc
    
    # Add multipliers if requested (do this BEFORE any merge_replace in the document)
    # Actually, merge_replace happens during parsing, so add_multipliers should happen
    # only to nodes that don't already have multipliers
    if add_multipliers:
        add_multipliers_recursive(merged_doc)
    
    return dump_yaml(merged_doc)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python deep_merge_yaml.py <yaml_file1> [yaml_file2 ...] [-o output_file] [--add-multipliers]", file=sys.stderr)
        sys.exit(1)

    args = sys.argv[1:]
    output_path = None
    input_files = []
    add_multipliers = False
    
    # Parse arguments
    i = 0
    while i < len(args):
        if args[i] == '-o' and i + 1 < len(args):
            output_path = Path(args[i + 1])
            i += 2
        elif args[i] == '--add-multipliers':
            add_multipliers = True
            i += 1
        else:
            input_files.append(Path(args[i]))
            i += 1
    
    if not input_files:
        print("Error: No input files specified", file=sys.stderr)
        sys.exit(1)

    try:
        merged_yaml = main(input_files, add_multipliers)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(2)

    if output_path:
        output_path.write_text(merged_yaml, encoding="utf-8")
        print(f"Output written to {output_path}")
        if add_multipliers:
            print("Multipliers added to all nodes with 'include' keys")
    else:
        print(merged_yaml)