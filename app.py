import os
import io
import re
import logging
import requests
from flask import (
    Flask, render_template, request, redirect, 
    url_for, flash, send_file, session, send_from_directory
)
from graphviz import Digraph

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.urandom(24)

# Create directories for outputs
os.makedirs('static/images', exist_ok=True)
os.makedirs('static/schemas', exist_ok=True)

# Constants
LISTS_TO_IGNORE = ["Documents", "Liens de partage", "Extensions de modèle web", "User"]
COLUMNS_TO_IGNORE = [
    "_ColorTag", "ComplianceAssetId", "_UIVersionString", "Attachments",
    "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "ItemChildCount",
    "FolderChildCount", "_ComplianceFlags", "_ComplianceTag", 
    "_ComplianceTagWrittenTime", "_ComplianceTagUserId", "_IsRecord",
    "AppAuthor", "AppEditor", "ID", "ContentType"
]
COLUMN_PATTERN_TO_IGNORE = ".*x003a.*"
GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0"

# Helper functions for API interactions
def create_headers(token):
    """Create headers for Graph API requests."""
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false;charset=utf-8",
        "Content-Type": "application/json"
    }

def fetch_data(url, headers):
    """Fetch data from the given URL with error handling."""
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logger.error(f"Request failed: {e}")
        return None

def get_column_type(column):
    """Determine the type of a column based on its properties."""
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

# SharePoint list and column operations
def fetch_sharepoint_lists(token, site_id):
    """Fetch all SharePoint lists excluding those in LISTS_TO_IGNORE."""
    endpoint = f"{GRAPH_API_BASE_URL}/sites/{site_id}/lists"
    headers = create_headers(token)
    
    lists_data = fetch_data(endpoint, headers)
    
    if not lists_data or 'value' not in lists_data:
        logger.error("Failed to fetch lists or no lists found.")
        return {}, headers, endpoint
    
    # Filter and process lists
    lists_dict = {}
    for item in lists_data['value']:
        if isinstance(item, dict) and 'displayName' in item and 'id' in item:
            if item['displayName'] not in LISTS_TO_IGNORE:
                lists_dict[item['displayName']] = item['id']
        else:
            logger.warning(f"Unexpected item format in 'value': {item}")
    
    return lists_dict, headers, endpoint

def fetch_columns(list_id, endpoint, headers):
    """Fetch and filter columns for a specific list."""
    current_endpoint = f"{endpoint}/{list_id}/columns"
    columns_data = fetch_data(current_endpoint, headers)
    
    if not columns_data or 'value' not in columns_data:
        logger.error(f"Failed to fetch columns for list '{list_id}'")
        return []
    
    # Filter and process columns
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

# Schema visualization functions
def generate_uml_graph(lists_dict, endpoint, headers):
    """Generate a UML graph of SharePoint lists and their relationships."""
    # Create a directed graph
    graph = Digraph(comment="Database Schema", format="png")
    graph.attr(rankdir="LR")  # Left-to-right layout

    # Fetch columns for each list and store relationships
    relationships = []
    
    for list_name, list_id in lists_dict.items():
        # Fetch columns
        columns = fetch_columns(list_id, endpoint, headers)
        logger.info(f"Generating table for {list_name}")
        
        # Create a label for the table
        label = f"<<TABLE BORDER='0' CELLBORDER='1' CELLSPACING='0'>\n"
        label += f"<TR><TD COLSPAN='2'><B>{list_name}</B></TD></TR>\n"
        
        for column in columns:
            column_name = column.get("name")
            column_type = column.get("type_details", {}).get("type", "unknown")
            label += f"<TR><TD>{column_name}</TD><TD>{column_type}</TD></TR>\n"
            
            # Check if this is a lookup column to establish relationships
            if column_type == "lookup":
                list_id_lookup = column.get("type_details", {}).get("details", {}).get("listId")
                if list_id_lookup:
                    relationships.append((list_name, list_id_lookup, column_name))
        
        label += "</TABLE>>"
        
        # Add the table as a node
        graph.node(list_name, label=label, shape="plaintext")

    # Add relationships (edges)
    for source_table, target_list_id, column_name in relationships:
        # Find the target table name
        target_table = None
        for table_name, table_id in lists_dict.items():
            if table_id == target_list_id:
                target_table = table_name
                break
        
        if target_table:
            # Add an edge from the source table to the target table
            graph.edge(source_table, target_table, label=column_name)

    # Return the graph as binary data
    return graph.pipe()

# Flask routes
@app.route('/', methods=['GET', 'POST'])
def index():
    """Handle the main page with the authentication form."""
    if request.method == 'POST':
        token = request.form['token']
        site_id = request.form['site_id']
       
        try:
            lists_dict, headers, endpoint = fetch_sharepoint_lists(token, site_id)
            
            if not lists_dict:
                flash('No SharePoint lists found or authentication failed.', 'error')
                return redirect(url_for('index'))
            
            # Store the data in session for later use
            session['lists_dict'] = lists_dict
            session['token'] = token
            session['site_id'] = site_id
            session['headers'] = headers
            session['endpoint'] = endpoint
            
            # Redirect to the results page
            return redirect(url_for('view_results'))
            
        except Exception as e:
            logger.error(f"Error processing request: {e}")
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    return render_template('index.html')

@app.route('/results')
def view_results():
    """Display the results page with the generated schema diagram."""
    if 'lists_dict' not in session:
        flash('Please submit the form first.', 'error')
        return redirect(url_for('index'))
    
    try:
        # Get the session data
        lists_dict = session['lists_dict']
        headers = session['headers']
        endpoint = session['endpoint']

        # Generate the UML graph
        graph_data = generate_uml_graph(lists_dict, endpoint, headers)
        
        # Save the graph image to a file
        graph_image_path = 'static/images/uml_graph.png'
        with open(graph_image_path, 'wb') as f:
            f.write(graph_data)
        
        # Pass the image path to the template
        return render_template('results.html', graph_path=graph_image_path)
    
    except Exception as e:
        logger.error(f"Error rendering results: {e}")
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/schema')
def view_schema():
    """Generate and return the schema diagram as an image."""
    if 'lists_dict' not in session:
        flash('Please submit the form first', 'error')
        return redirect(url_for('index'))
    
    # Generate the UML graph using the stored data
    lists_dict = session['lists_dict']
    endpoint = session['endpoint']
    headers = session['headers']
    image_data = generate_uml_graph(lists_dict, endpoint, headers)
    
    # Return the image data
    return send_file(
        io.BytesIO(image_data),
        mimetype='image/png'
    )

@app.route('/download/<filename>')
def download_file(filename):
    """Handle file downloads."""
    try:
        return send_from_directory('static/images', filename)
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        flash('Error downloading the file.', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)