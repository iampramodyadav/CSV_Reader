"""
Specific Helpers
Component graphs and column suggestions
"""
import os
import re
import pandas as pd
from pyvis.network import Network
from config import GRAPH_TYPE_COLORS, MDT_DROPDOWNS
import numpy as np
from typing import Iterable, Optional, Union

def plot_interactive_component_graph(csv_file: str, output_file: str = "graph.html",
                                     open_browser: bool = True, comp_csv: str = None):
    """
    Build an interactive PyVis graph:
    - CSV1 must have columns: COMP1, COMP2, Type
    - CSV2 (optional) may have column: Name or COMP (list of allowed component names)
    - Multiple identical edges display their count
    """
    # Primary CSV
    df = pd.read_csv(csv_file, sep=",", engine="python").fillna('Unconnected')
    edge_counts = df.groupby(["COMP1", "COMP2", "Type"]).size().reset_index(name="count")

    # Allowed nodes from secondary CSV (if provided)
    allowed_nodes = None
    if comp_csv and os.path.exists(comp_csv):
        comp_df = pd.read_csv(comp_csv, sep=",", engine="python")
        name_col = "Name" if "Name" in comp_df.columns else ("COMP" if "COMP" in comp_df.columns else None)
        if name_col:
            allowed_nodes = set(comp_df[name_col].dropna().astype(str).tolist())

    # PyVis setup
    net = Network(height="800px", width="100%", directed=True)

    # Colors per type (fallback to gray)
    type_colors = {
        0: "gray",
        1: "purple",
        2: "blue",
        3: "cyan",
        4: "green",
        5: "red",
        6: "orange",
    }
    type_color_used = {}

    # Add nodes and edges
    for _, row in edge_counts.iterrows():
        u, v, count, typec = row["COMP1"], row["COMP2"], row["count"], str(row["Type"])

        def node_color(node: str) -> str:
            if node == "Unconnected":
                return "gray"
            if allowed_nodes is not None and node not in allowed_nodes:
                return "gray"
            return "#4CAF50"

        net.add_node(u, label=u, color=node_color(u), size=25)
        net.add_node(v, label=v, color=node_color(v), size=25)

        try:
            edge_color = type_colors.get(int(float(typec)), "#AAAAAA")
        except Exception:
            edge_color = "#AAAAAA"
        type_color_used[typec] = edge_color

        edge_options = {"color": edge_color}
        if count > 1:
            edge_options.update({
                "label": str(count),
                "font": {"color": "blue", "size": 14, "strokeWidth": 1},
                "width": 2
            })
        net.add_edge(u, v, **edge_options)

    # Legend
    i = 0
    for typec, color in type_color_used.items():
        legend_node = f"Legend: Type {typec}"
        net.add_node(
            legend_node,
            label=f"Conta Type {typec}",
            color=color,
            shape='box',
            size=20,
            physics=False,
            fixed=True,
            x=-850,
            y=0 + 50 * i
        )
        i += 1
    if allowed_nodes is not None:
        net.add_node(
            "Legend: Missing",
            label="Not in comp name",
            color="gray",
            shape="box",
            size=20,
            physics=False,
            fixed=True,
            x=-850,
            y=0 + 50 * i
        )

    net.write_html(output_file, open_browser=open_browser)


def plot_interactive_component_graph_coordinate(df, output_file: str = "graph.html",
                                     open_browser: bool = True, comp_csv: str = None):
    """
    Build an interactive PyVis graph from interface names.
    - CSV format: Single column with values like "COMP1_COMP2_INTF<number>"
    - Extracts COMP1, COMP2 from the interface name
    - CSV2 (optional) may have column: Name or COMP (list of allowed component names)
    - Multiple identical edges display their count
    """
    # Get the first column name
    first_col = df.columns[0]
    
    # Parse interface names to extract COMP1 and COMP2
    parsed_data = []
    for intf_name in df[first_col]:
        # Split by underscore
        parts = str(intf_name).split('_')
        
        if len(parts) >= 3:
            # Find the part that starts with "INTF"
            intf_idx = None
            for i, part in enumerate(parts):
                if part.startswith('INTF'):
                    intf_idx = i
                    break
            
            if intf_idx is not None and intf_idx >= 2:
                # COMP1 is everything before COMP2
                comp1 = '_'.join(parts[:intf_idx-1])
                # COMP2 is the part just before INTF
                comp2 = parts[intf_idx-1]
                
                parsed_data.append({
                    'COMP1': comp1,
                    'COMP2': comp2
                })
    
    # Create DataFrame from parsed data
    df_parsed = pd.DataFrame(parsed_data)
    
    # Count edges
    edge_counts = df_parsed.groupby(["COMP1", "COMP2"]).size().reset_index(name="count")

    # Allowed nodes from secondary CSV (if provided)
    allowed_nodes = None
    if comp_csv and os.path.exists(comp_csv):
        comp_df = pd.read_csv(comp_csv, sep=",", engine="python")
        name_col = "Name" if "Name" in comp_df.columns else ("COMP" if "COMP" in comp_df.columns else None)
        if name_col:
            allowed_nodes = set(comp_df[name_col].dropna().astype(str).tolist())

    # PyVis setup
    net = Network(height="800px", width="100%", directed=True)

    # Add nodes and edges
    for _, row in edge_counts.iterrows():
        u, v, count = row["COMP1"], row["COMP2"], row["count"]

        def node_color(node: str) -> str:
            if allowed_nodes is not None and node not in allowed_nodes:
                return "gray"
            return "#4CAF50"

        net.add_node(u, label=u, color=node_color(u), size=25)
        net.add_node(v, label=v, color=node_color(v), size=25)

        edge_options = {"color": "#2196F3"}
        if count > 1:
            edge_options.update({
                "label": str(count),
                "font": {"color": "blue", "size": 14, "strokeWidth": 1},
                "width": 2
            })
        net.add_edge(u, v, **edge_options)

    # Legend
    if allowed_nodes is not None:
        net.add_node(
            "Legend: Missing",
            label="Not in comp list",
            color="gray",
            shape="box",
            size=20,
            physics=False,
            fixed=True,
            x=-850,
            y=0
        )

    net.write_html(output_file, open_browser=open_browser)
    print(f"Graph saved to {output_file}")
    print(f"Parsed {len(parsed_data)} interfaces into {len(edge_counts)} unique edges")

def get_mdt_dropdown_config(filename: str) -> dict:
    """
    Get dropdown configuration based on filename prefix.
    Returns dict of {column_name: options}
    """
    if filename.startswith("Turbine_Comp"):
        return MDT_DROPDOWNS["Turbine_Comp"]
    elif filename.startswith("Turbine_Conta"):
        return MDT_DROPDOWNS["Turbine_Conta"]
    elif filename.startswith("Turbine_Joint"):
        return MDT_DROPDOWNS["Turbine_Joint"]
    return {}


def get_file_list_for_column(folder_path: str, column_name: str) -> list:
    """
    Get file list for specific columns (e.g., Parasolid column needs .x_t files).
    """
    if column_name == "Parasolid":
        files = [f for f in os.listdir(folder_path) if f.endswith(".x_t") or f.endswith(".cdb")]
        files.append('')
        return files
    return []

def trim_whitespace(
    df: pd.DataFrame,
    columns: Optional[Iterable[str]] = None
) -> pd.DataFrame:
    """
    Trim leading/trailing whitespace in string-like columns.
    If `columns` is None, only object/StringDtype columns are trimmed.
    """
    out = df.copy()
    if columns is None:
        cols = out.select_dtypes(include=["object", "string"]).columns
    else:
        cols = list(columns)

    for col in cols:
        out[col] = out[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return out


def fill_nulls(
    df: pd.DataFrame,
    columns: Optional[Iterable[str]] = None,
    fill_value: Union[str, int, float] = "",
    only_object_columns: bool = True,
    strip_whitespace: bool = True
) -> pd.DataFrame:
    """
    More aggressive version that also strips whitespace before checking for null-like values.
    
    Additional Parameters:
    ----------------------
    strip_whitespace : bool
        If True, strips leading/trailing whitespace before comparison
    """
    out = df.copy()
    
    if columns is None:
        if only_object_columns:
            cols = out.select_dtypes(include=["object", "string"]).columns
        else:
            cols = out.columns
    else:
        cols = list(columns)

    null_like_values = [
        np.nan, None, 'NaN', 'nan', 'None', 'null', 'NULL', 
        'Null', '<NA>', 'NA', 'N/A', 'n/a'
    ]
    
    for col in cols:
        # Replace actual NaN/None
        out[col] = out[col].replace({np.nan: fill_value, None: fill_value})
        
        if out[col].dtype == 'object' or pd.api.types.is_string_dtype(out[col]):
            # Strip whitespace if requested
            if strip_whitespace:
                out[col] = out[col].astype(str).str.strip()
            
            # Replace string representations (case-insensitive)
            out[col] = out[col].replace(null_like_values, fill_value)
            
            # Replace empty strings and whitespace-only
            out[col] = out[col].replace(r'^\s*$', fill_value, regex=True)
            
            # If we stripped and created empty strings, replace them too
            if strip_whitespace:
                out[col] = out[col].replace('', fill_value)
    
    return out

def drop_fully_empty_rows(
    df: pd.DataFrame,
    consider_columns: Optional[Iterable[str]] = None
) -> pd.DataFrame:
    """
    Drop rows where *all* considered columns are empty (NaN/None or empty strings).
    If `consider_columns` is None, use all columns.
    """
    out = df.copy()
    cols = list(consider_columns) if consider_columns is not None else out.columns

    # Treat empty strings as NaN temporarily for the check
    temp = out[cols].replace("", np.nan)
    mask_all_empty = temp.isna().all(axis=1)
    return out.loc[~mask_all_empty].copy()

def read_label_format(path):
    """
    Reads a label-format data file and returns:
      - labels list
      - identifiers list (same as labels)
      - data list (list of lists)
      - pandas DataFrame with all columns
    """
    labels = []
    data = []

    with open(path, "r") as f:
        lines = [ln.rstrip() for ln in f.readlines() if ln.strip()]

    # Parse header
    header = re.split(r"\s{2,}", lines[0].strip())
    header[0] = "Label"
    identifiers =header[1:]
    
    # Parse data rows
    for ln in lines[1:]:
        if "::" not in ln:
            continue

        label, raw = ln.split("::")
        label = label.strip()
        labels.append(label)
        # identifiers.append(label)

        # Split numeric columns
        values = re.split(r"\s+", raw.strip())

        # Convert numbers
        parsed = []
        for v in values:
            try:
                if "." in v:
                    parsed.append(float(v))
                else:
                    parsed.append(int(v))
            except:
                parsed.append(v)

        data.append(parsed)

    # Build DataFrame
    df = pd.DataFrame(data, columns=header[1:])
    df.insert(0, "#label", labels)

    return labels, identifiers, data, df

def write_label_format(path, df, sep="::"):
    """
    Writes aligned label-format file from DataFrame.
    """
    df = df.rename(columns={df.columns[0]: '#label'})
    header = list(df.columns)
    
    # Compute column widths
    col_widths = {h: max(len(h), df[h].astype(str).map(len).max()) for h in header}

    with open(path, "w") as f:

        # HEADER
        header_line = header[0].ljust(col_widths[header[0]]+5)
        for h in header[1:]:
            header_line += "  " + h.rjust(col_widths[h])
        f.write(header_line + "\n")

        # DATA LINES
        for _, row in df.iterrows():
            line = str(row[header[0]]).ljust(col_widths[header[0]])
            line += f"   {sep}"

            for h in header[1:]:
                line += "  " + str(row[h]).rjust(col_widths[h])
            f.write(line + "\n")


# if __name__ == "__main__":
#     path = r'coordinateTable_FSTMKVII.txt'
#     path1 = r'coordinateTable_FSTMKVII2.txt'
#     comp_path = r"K:\Groups\OFTDDDMKVII\BEA\SFFB\FST\00\Turbine_Comp_Catalogue_FSTMKVIIGR00GI00.csv"
#     coordinateTable = r"K:\Groups\OFTDDDMKVII\BEA\SFFB\FST\00\coordinateTable_FSTMKVII.txt"
#     labels, identifiers, data, df = read_label_format(coordinateTable)
#     write_label_format(path1, df, sep="::")
#     plot_interactive_component_graph_coordinate(df,comp_csv =comp_path  )