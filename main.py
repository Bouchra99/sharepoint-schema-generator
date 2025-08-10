import os
import re
import io
import logging
import argparse
import requests
from graphviz import Digraph

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

os.makedirs('graph', exist_ok=True)

LISTS_TO_IGNORE = ["Documents", "Liens de partage", "Extensions de modèle web", "User", "Web Template Extensions"]
COLUMNS_TO_IGNORE = [
    "_ColorTag", "ComplianceAssetId", "_UIVersionString", "Attachments",
    "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "ItemChildCount",
    "FolderChildCount", "_ComplianceFlags", "_ComplianceTag", 
    "_ComplianceTagWrittenTime", "_ComplianceTagUserId", "_IsRecord",
    "AppAuthor", "AppEditor", "ID", "ContentType"
]
COLUMN_PATTERN_TO_IGNORE = ".*x003a.*"
GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0"

def create_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false;charset=utf-8",
        "Content-Type": "application/json"
    }

def fetch_data(url, headers):
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logger.error(f"Request failed: {e}")
        return None

def get_column_type(column):
    
    type_mappings = {
        "text": "text",
        "lookup": "lookup",
        "dateTime": "dateTime",
        "number": "number",
        "choice": "choice",
        "boolean": "boolean",
        "person": "person",
        "calculated": "calculated",
    }
    
    for key, value in type_mappings.items():
        if key in column:
            return {"type": value, "details": column[key]}
    
    return {"type": "unknown", "details": {}}

def fetch_sharepoint_lists(token, site_id):
    """Fetch all SharePoint lists except LISTS_TO_IGNORE  """
    endpoint = f"{GRAPH_API_BASE_URL}/sites/{site_id}/lists"
    headers = create_headers(token)
    
    lists_data = fetch_data(endpoint, headers)
    
    if not lists_data or 'value' not in lists_data:
        logger.error("Failed to fetch lists or no lists found.")
        return {}, headers, endpoint
    
    lists_dict = {}
    for item in lists_data['value']:
        if isinstance(item, dict) and 'displayName' in item and 'id' in item:
            if item['displayName'] not in LISTS_TO_IGNORE:
                lists_dict[item['displayName']] = item['id']
        else:
            logger.warning(f"Unexpected item format in 'value': {item}")
    
    return lists_dict, headers, endpoint

def fetch_columns(list_id, endpoint, headers):
    """Fetch  columns for a specific list"""
    current_endpoint = f"{endpoint}/{list_id}/columns"
    columns_data = fetch_data(current_endpoint, headers)
    
    if not columns_data or 'value' not in columns_data:
        logger.error(f"Failed to fetch columns for list '{list_id}'")
        return []
    
    columns = []
    for col in columns_data['value']:
        name = col.get("name", "")
        if (name not in COLUMNS_TO_IGNORE and 
            not re.match(COLUMN_PATTERN_TO_IGNORE, name)):
            columns.append({
                "name": name,
                "id": col.get("id"),
                "required": col.get("required", False),
                "type_details": get_column_type(col)
            })
    
    return columns

def generate_uml_graph(lists_dict, endpoint, headers):
    graph = Digraph(comment="Database Schema", format="png")
    graph.attr(rankdir="LR")

    relationships = []
    
    for list_name, list_id in lists_dict.items():
        columns = fetch_columns(list_id, endpoint, headers)
        logger.info(f"Generating table for {list_name}")
        
        label = f"<<TABLE BORDER='0' CELLBORDER='1' CELLSPACING='0'>\n"
        label += f"<TR><TD COLSPAN='2'><B>{list_name}</B></TD></TR>\n"
        
        for column in columns:
            column_name = column.get("name")
            column_type = column.get("type_details", {}).get("type", "unknown")
            label += f"<TR><TD>{column_name}</TD><TD>{column_type}</TD></TR>\n"
            
            if column_type == "lookup":
                list_id_lookup = column.get("type_details", {}).get("details", {}).get("listId")
                if list_id_lookup:
                    relationships.append((list_name, list_id_lookup, column_name))
        
        label += "</TABLE>>"
        graph.node(list_name, label=label, shape="plaintext")

    for source_table, target_list_id, column_name in relationships:
        target_table = next((name for name, id_ in lists_dict.items() if id_ == target_list_id), None)
        if target_table:
            graph.edge(source_table, target_table, label=column_name)

    return graph.pipe()

def main(token, site_id):
    lists_dict, headers, endpoint = fetch_sharepoint_lists(token, site_id)
    
    if not lists_dict:
        logger.error("No SharePoint lists found or authentication failed")
        return
    
    graph_data = generate_uml_graph(lists_dict, endpoint, headers)
    
    output_path = os.path.join("graph", "uml_graph.png")
    with open(output_path, "wb") as f:
        f.write(graph_data)
    
    logger.info(f"✅ UML diagram saved to {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate UML diagram from SharePoint site")
    parser.add_argument("--token", required=True, help="Access token for Microsoft Graph API")
    parser.add_argument("--site-id", required=True, help="SharePoint site ID")
    args = parser.parse_args()

    main(args.token, args.site_id)
