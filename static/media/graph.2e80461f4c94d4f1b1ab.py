"""
Organizational Chart Generator
-----------------------------
This script processes organizational data from Excel files and generates visual organizational charts as PNG images.

Main Features:
- Loads, cleans, and processes org data for a specified manager and month/year.
- Handles duplicate employee names and missing data.
- Maps job titles to grades using a lookup table.
- Generates hierarchical org charts using Graphviz and saves them as images.
- Optionally adds branding and department labels to the final chart image.

Key Dependencies:
- pandas: For data manipulation and cleaning.
- graphviz: For generating org chart visualizations.
- Pillow (PIL): For image processing and branding overlays.

Usage:
- Can be run as a standalone CLI tool (see main() and parse_args()).
- Can be integrated with a GUI frontend for interactive use.

Directory Structure:
- Expects input Excel files in the '../data/' directory.
- Outputs cleaned data to '../all_reporting/'.
- Outputs raw org chart images to '../raw_graph/'.
- Outputs finalized org chart images to '../org chart/'.

"""
import pandas as pd
import os
import argparse
from graphviz import Digraph
from PIL import Image, ImageDraw, ImageFont

# Ensure Graphviz is available in the system PATH
os.environ["PATH"] += os.pathsep + 'C:/Program Files/Graphviz/bin'

def load_data(manager_name, month_year, file_path=None):
    """
    Load and preprocess the Excel data for a given manager and month/year, or from a direct file path.
    - If file_path is provided, loads from that path.
    - Otherwise, loads from the default location using manager_name and month_year.
    - Renames columns for consistency.
    - Cleans up the 'Reports To' field and column names.
    
    Args:
        manager_name (str): Name of the manager (used in file naming).
        month_year (str): Month and year string (used in file naming).
        file_path (str, optional): Direct path to the Excel file. Defaults to None.
    
    Returns:
        pd.DataFrame: Preprocessed DataFrame containing org data.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if file_path is not None:
        input_path = file_path
    else:
        input_path = os.path.join(script_dir, '..', 'data', f'{manager_name} {month_year}.xlsx')
    df = pd.read_excel(input_path, sheet_name=0, engine='openpyxl')
    df.rename(columns={'Line Detail 1': 'Title', 'Line Detail 2': 'Location'}, inplace=True, errors='raise')
    df['Reports To'] = df['Reports To'].apply(lambda x: ' '.join(x.split('_'))[1:] if isinstance(x, str) else x)
    df.columns = df.columns.str.upper().str.strip().str.replace(' ', '_')
    return df

def save_df(datasheet, manager_name, month_year):
    """
    Clean and save the DataFrame to the all_reporting directory.
    - Removes duplicate columns.
    - Selects and renames columns for consistency.
    - Fills missing values and sets status.
    - Maps titles to grades using a lookup table.
    - Handles duplicate employees.
    - Saves the cleaned DataFrame to an Excel file.
    
    Args:
        datasheet (pd.DataFrame): Raw org data.
        manager_name (str): Name of the manager (used in file naming).
        month_year (str): Month and year string (used in file naming).
    
    Returns:
        pd.DataFrame: Cleaned DataFrame ready for graph generation.
    """
    df = datasheet
    remove_duplicates = df.loc[:, ~df.columns.duplicated(keep='last')]
    selected_columns = (remove_duplicates[['NAME', 'REPORTS_TO', 'TITLE', 'LOCATION']])
    selected_columns.columns = selected_columns.columns.str.replace(' ', '_')
    renamed_columns = selected_columns.rename(columns={'NAME': 'FULL_NAME', 'REPORTS_TO': 'MANAGER', 'LEVEL': 'GRADE'})
    manager_name = str(manager_name).replace(' ', '_')
    df = renamed_columns.copy()
    df.loc[:, 'STATUS'] = 'ACTIVE'
    empty_row_index = df['TITLE'].isna() | df['TITLE'].eq('')
    df.loc[empty_row_index, 'TITLE'] = df.loc[empty_row_index, 'FULL NAME']
    df.loc[empty_row_index, 'FULL NAME'] = 'VACANT'
    df.loc[empty_row_index, 'STATUS'] = 'VACANT'

    # Map titles to grades
    title_to_level_dict = load_title_level_dict()
    title_to_level_dict = dict(zip(title_to_level_dict['TITLE'], title_to_level_dict['GRADE']))
    df.loc[:, 'GRADE'] = df['TITLE'].map(title_to_level_dict)
    df.loc[:, 'GRADE'] = df['GRADE'].fillna('')

    # Mark employees on leave
    status_row_index = df['FULL NAME'].str.contains(r' \(On Leave\)', case=False, na=False)
    df.loc[status_row_index, 'STATUS'] = 'ON_LEAVE'
    df['FULL NAME'] = df['FULL NAME'].str.replace(r' \(On Leave\)', '', regex=True).str.strip()
    df['FULL NAME'] = df['FULL NAME'].str.replace(r' \[C\]', '', regex=True).str.strip()
    df['TITLE'] = df['TITLE'].str.replace(r' \(Unfilled\)', '', regex=True).str.strip()

    # For visualization, use a unique name if duplicate
    df['EMPID'] = df['FULL NAME']

    df = identify_duplicate_employees(df)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, '..', 'all_reporting', f'{manager_name} {month_year}.xlsx')
    if not os.path.exists(output_path):
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=month_year, index=False)
    else:
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=month_year, index=False)
    return df

def identify_duplicate_employees(df):
    """
    Identify and mark duplicate employee names in the DataFrame for visualization.
    - If an employee appears more than once, append a unique identifier to their name.
    
    Args:
        df (pd.DataFrame): DataFrame with employee data.
    
    Returns:
        pd.DataFrame: DataFrame with unique 'EMPID' for duplicates.
    """
    duplicate_employees = df['FULL NAME'][df['FULL NAME'].duplicated(keep=False)].unique()
    for name in duplicate_employees:
        employee = df[df['FULL NAME'] == name]
        if employee['MANAGER'].nunique():
            marked_name = [f'{name} ({i+1})' for i in range(len(employee))]
        else:
            marked_name = employee['MANAGER'].apply(lambda x: f'{name} ({x.split()[-1]})')
        df.loc[employee.index, 'EMPID'] = marked_name
    return df

def load_title_level_dict():
    """
    Load the title-to-grade mapping from an Excel file.
    
    Returns:
        pd.DataFrame: DataFrame with 'TITLE' and 'GRADE' columns.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(script_dir, '..', 'data', 'title_level_dict.xlsx')
    df = pd.read_excel(input_path, sheet_name=0, engine='openpyxl')
    return df

def build_dict(reports_by_manager, root):
    """
    Build a nested dictionary representing the reporting structure starting from the root manager.
    
    Args:
        reports_by_manager (dict): Mapping from manager to list of direct reports.
        root (str): The root manager's name.
    
    Returns:
        dict: Nested dictionary representing the org structure.
    """
    tree = {}

    def dfs(node):
        children = reports_by_manager.get(node, [])
        if children:
            tree[node] = children
            for child in children:
                dfs(child)
        return tree

    return dfs(root)

def generateGraph(df, manager_name, month_year, show_location=False, show_level=False):
    """
    Generate and save the org chart as a PNG image using Graphviz.
    - Builds the reporting structure and visualizes it.
    - Supports options to show location and level in node labels.
    
    Args:
        df (pd.DataFrame): Cleaned org data.
        manager_name (str): Name of the root manager.
        month_year (str): Month and year string for labeling.
        show_location (bool): Whether to display location in node labels.
        show_level (bool): Whether to display level/grade in node labels.
    
    Returns:
        Digraph: The generated Graphviz Digraph object.
    """
    # Precompute mappings and label builder
    df_by_name = df.set_index('FULL NAME', drop=False)
    reports_by_manager = df.groupby('MANAGER')['FULL NAME'].apply(list).to_dict()

    def make_label(emp_name):
        """Build the label for a node, including optional fields."""
        row = df_by_name.loc[emp_name]
        parts = [str(row['FULL NAME']), str(row['TITLE'])]
        if show_level:
            parts.append(f"Level: {row['LEVEL']}")
        if show_location:
            parts.append(str(row['LOCATION']))
        return "\n".join(parts)

    dot = Digraph()
    dot.attr(rankdir='TB', splines='ortho', nodesep='0.6', ranksep='0.8')
    dot.attr('edge', arrowhead='none', color='black', penwidth='1.5')
    dot.attr('node',
             shape='box',
             style='filled',
             fillcolor='#F5F5F5',
             width='2.5',
             height='1.5',
             fontsize='14',
             penwidth='1')

    # Add the root manager node
    directing_manager_label = make_label(manager_name)
    dot.node(manager_name, label=directing_manager_label, peripheries='2')
    direct_reports = reports_by_manager.get(manager_name, [])

    # Add direct reports as a subgraph
    with dot.subgraph() as direp:
        for emp in direct_reports:
            emp_label = make_label(emp)
            direp.node(emp, label=emp_label)
            direp.edge(manager_name, emp)

    # Build subteams recursively
    subteams = build_dict(reports_by_manager, manager_name)
    subteams.pop(manager_name, None)

    for man, reports in subteams.items():
        with dot.subgraph(name=f'cluster_{man}') as sub:
            sub.attr(style='invis')
            man_label = make_label(man)
            sub.node(man, label=man_label)
            n = len(reports)

            if n<=3:
                sub.attr(rank='same')
                for emp in reports:
                    emp_label = make_label(emp)
                    sub.node(emp, label=emp_label)
                    sub.edge(man, emp)
            else:
                # For large teams, use bus layout for clarity
                bus_node = [f'bus_{man}_{i}' for i in range(n)]
                for bn in bus_node:
                    sub.node(bn, label='', shape='point', width='0.01', height='0.01')
                sub.edge(man, bus_node[0], arrowhead='none', weight='10')

                for i in range(len(bus_node)-1):
                    sub.edge(bus_node[i], bus_node[i+1], arrowhead='none', weight='5')

                for i, emp in enumerate(reports):
                    sub.node(emp, label=make_label(emp))
                    sub.edge(bus_node[i], emp, constraint='false', minlen='1', weight='1', arrowhead='none')

                for i in range(n):
                    with dot.subgraph() as same_rank:
                        same_rank.attr(rank='same')
                        same_rank.node(bus_node[i])
                        emp_label = make_label(reports[i])
                        same_rank.node(reports[i], label=emp_label)

    # Save the org chart as a PNG image
    org_chart_name = f'{manager_name} {month_year}'
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, '..', 'raw_graph', f'{org_chart_name}.png')
    dot.render(output_path, format='png', cleanup=True)
    return org_chart_name

def finalize_graph(org_chart_name, month_year):
    """
    Add branding and labels to the generated org chart image.
    - Loads the raw org chart PNG and overlays logo and department labels.
    - Saves the final image to the org chart directory.
    
    Args:
        org_chart_name (str): Name of the org chart (usually manager + month_year).
        month_year (str): Month and year string for labeling.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    graph_path = os.path.join(script_dir, '..', 'raw_graph', f'{org_chart_name} Graph.png')
    output_path = os.path.join(script_dir, '..', 'org chart', f'{org_chart_name} Org Chart.png')
    logo_path = os.path.join(script_dir, '..', 'data', 'logo.png')
    graph = Image.open(graph_path)
    logo = Image.open(logo_path)
    logo.thumbnail((384, 384))
    label1 = "Global Trade Finance Operations"
    label2 = "Global Operations & Business Services"
    font = ImageFont.truetype("arial.ttf", 40)
    font1 = ImageFont.truetype("arialbd.ttf", 42)

    width, height = graph.size
    logo_width, logo_height = logo.size
    text_shift = logo_width + 200

    padded_left_right = 1000
    padded_top_bottom = 400

    padded_width = width + padded_left_right*2
    padded_height = height + padded_top_bottom*2

    padded_graph = Image.new(graph.mode, (padded_width, padded_height), (255, 255, 255))
    padded_graph.paste(graph, (padded_left_right, padded_top_bottom))
    padded_graph.paste(logo, (200, 200))
    draw = ImageDraw.Draw(padded_graph)
    draw.text((text_shift, 200), label1, font=font1, fill=(0, 0, 0))
    draw.text((text_shift, 250), label2, font=font, fill=(0, 0, 0))
    draw.text((text_shift, 300), month_year, font=font, fill=(0, 0, 0))
    padded_graph.save(fp=output_path, format='PNG')


def parse_args():
    """
    Parse command-line arguments for the script.
    Returns:
        argparse.Namespace: Parsed arguments with manager and month_year.
    """
    parser = argparse.ArgumentParser(description='Generate org charts from Excel data')
    parser.add_argument('manager', type=str, required=True, help='Name of the manager to generate chart for')
    parser.add_argument('month_year', type=str, required=True, help='Month and year of the data')
    return parser.parse_args()

def main():
    """
    Main entry point for the script when run as a standalone program.
    - Parses arguments.
    - Loads and cleans data.
    - Generates the org chart.
    """
    args = parse_args()
    manager_name = args.manager
    month_year = args.month_year

    datasheet = load_data(manager_name, month_year)
    df = save_df(datasheet, manager_name, month_year)

    graph_name = generateGraph(df, manager_name, month_year)
    finalize_graph(graph_name, month_year)

if __name__ == '__main__':
    main()