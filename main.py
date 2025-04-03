import pandas as pd
import numpy as np
import os
import json
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
# Add this at the top of your app.py file, right after your imports
import streamlit as st
import time
import datetime
import hashlib

# Set page configuration
st.set_page_config(
    page_title="Office Room Allocation System",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom CSS
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton button {
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
    }
    .stAlert {
        border-radius: 10px;
    }
    .room-vacant {
        background-color: #d4edda;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 5px solid #28a745;
    }
    .room-low {
        background-color: #e6f7e1;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 5px solid #5cb85c;
    }
    .room-medium {
        background-color: #fff3cd;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 5px solid #ffc107;
    }
    .room-high {
        background-color: #ffe5d9;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 5px solid #fd7e14;
    }
    .room-full {
        background-color: #f8d7da;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 5px solid #dc3545;
    }
    .room-storage {
        background-color: #e2e3e5;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 5px solid #6c757d;
    }
    .occupant-tag {
        display: inline-block;
        background-color: #e9ecef;
        padding: 2px 8px;
        border-radius: 12px;
        margin: 2px;
        font-size: 0.8em;
    }
    .status-current {
        background-color: #d4edda;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: bold;
        color: #155724;
    }
    .status-upcoming {
        background-color: #cce5ff;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: bold;
        color: #004085;
    }
    .status-past {
        background-color: #f8d7da;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: bold;
        color: #721c24;
    }
    .floor-heading {
        background-color: #f1f3f5;
        padding: 5px 10px;
        border-radius: 5px;
        margin-top: 15px;
        margin-bottom: 10px;
    }
    .dataframe th {
        font-size: 14px;
        font-weight: bold;
        background-color: #f1f3f5;
    }
    .dataframe td {
        font-size: 13px;
    }
    </style>
    """, unsafe_allow_html=True)

# Function to check password
def check_password(password):
    # In a real application, you would use a more secure approach
    # For this example, we're using a simple hash comparison
    correct_password_hash = hashlib.sha256("kluth2025".encode()).hexdigest()
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    return password_hash == correct_password_hash

# Function to check if session has timed out
def check_session_timeout():
    # Get the current time
    current_time = datetime.datetime.now()
    
    # Check if login_time exists in session state
    if 'login_time' in st.session_state:
        # Calculate time difference in hours
        time_diff = current_time - st.session_state.login_time
        hours_diff = time_diff.total_seconds() / 3600
        
        # If more than 1 hour has passed, session has timed out
        if hours_diff > 1:
            st.session_state.is_authenticated = False
            st.session_state.pop('login_time', None)
            return True
    
    return False

# Authentication function
def authenticate():
    # Check if already authenticated and session hasn't timed out
    if 'is_authenticated' in st.session_state and st.session_state.is_authenticated:
        # Check if session has timed out
        if check_session_timeout():
            st.warning("Your session has timed out. Please log in again.")
            # Show login form again
            show_login_form()
        else:
            # User is authenticated and session is still valid
            return True
    else:
        # User is not authenticated, show login form
        show_login_form()
        return False

# Function to display login form
def show_login_form():
    st.title("Office Room Allocation System - Login")
    st.markdown("Please enter the password to access the system.")
    
    # Create login form
    with st.form("login_form"):
        password = st.text_input("Password", type="password")
        submit_button = st.form_submit_button("Login")
        
        if submit_button:
            if check_password(password):
                # Set authentication status and login time
                st.session_state.is_authenticated = True
                st.session_state.login_time = datetime.datetime.now()
                st.success("Login successful!")
                # Add a rerun to refresh the page after successful login
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")

# Add this right after your imports and authentication functions
# but before the main application code
if not authenticate():
    # If not authenticated, stop here
    st.stop()

# Display remaining session time if authenticated
if 'login_time' in st.session_state:
    current_time = datetime.datetime.now()
    time_diff = current_time - st.session_state.login_time
    remaining_minutes = max(0, 60 - (time_diff.total_seconds() / 60))
    
    # Display a message when less than 10 minutes are remaining
    if remaining_minutes < 10:
        st.sidebar.warning(f"⚠️ Session expires in {int(remaining_minutes)} minutes")
    else:
        st.sidebar.info(f"Session time remaining: {int(remaining_minutes)} minutes")

# Add a logout button in the sidebar
if st.sidebar.button("Logout"):
    # Clear authentication status
    st.session_state.is_authenticated = False
    if 'login_time' in st.session_state:
        del st.session_state.login_time
    st.sidebar.success("Logged out successfully!")
    # Rerun the app to show login form
    st.rerun()



# Create directories if they don't exist
os.makedirs('data', exist_ok=True)
os.makedirs('data/backup', exist_ok=True)

# File paths
DEFAULT_EXCEL_PATH = 'data/MP_Office_Allocation.xlsx'
CAPACITY_CONFIG_PATH = 'data/room_capacities.json'

# Initialize session state
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.current_df = pd.DataFrame()
    st.session_state.upcoming_df = pd.DataFrame()
    st.session_state.past_df = pd.DataFrame()
    st.session_state.file_path = DEFAULT_EXCEL_PATH
    st.session_state.filter_building = 'All'
    st.session_state.last_save = None
    st.session_state.room_capacities = {}

# Function to load room capacities
def load_room_capacities():
    try:
        if os.path.exists(CAPACITY_CONFIG_PATH):
            with open(CAPACITY_CONFIG_PATH, 'r') as f:
                capacities = json.load(f)
            return capacities
        return {}
    except Exception as e:
        st.error(f"Error loading room capacities: {e}")
        return {}

# Function to save room capacities
def save_room_capacities(capacities):
    try:
        with open(CAPACITY_CONFIG_PATH, 'w') as f:
            json.dump(capacities, f)
        return True
    except Exception as e:
        st.error(f"Error saving room capacities: {e}")
        return False

# Load room capacities if not already loaded
if not st.session_state.room_capacities:
    st.session_state.room_capacities = load_room_capacities()

# Function to load data from Excel file
@st.cache_data(ttl=300)
def load_data(file_path=DEFAULT_EXCEL_PATH):
    try:
        # Load all sheets
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        # Identify sheets
        current_sheet = next((s for s in sheet_names if 'current' in s.lower()), sheet_names[0] if sheet_names else None)
        upcoming_sheet = next((s for s in sheet_names if 'upcoming' in s.lower()), None)
        past_sheet = next((s for s in sheet_names if 'past' in s.lower()), None)
        
        # Load sheets into dataframes
        current_df = pd.read_excel(file_path, sheet_name=current_sheet) if current_sheet else pd.DataFrame()
        upcoming_df = pd.read_excel(file_path, sheet_name=upcoming_sheet) if upcoming_sheet else pd.DataFrame()
        past_df = pd.read_excel(file_path, sheet_name=past_sheet) if past_sheet else pd.DataFrame()
        
        # Clean column names by stripping whitespace
        for df in [current_df, upcoming_df, past_df]:
            if not df.empty:
                df.columns = df.columns.str.strip()
        
        # Standardize column names
        column_mapping = {
            'Office': 'Office',
            'Office    ': 'Office',
            'Room': 'Office',
            'Room Number': 'Office',
            'Building': 'Building',
            'Building    ': 'Building',
            'Location': 'Building',
            'Email': 'Email address'
        }
        
        for df in [current_df, upcoming_df, past_df]:
            if not df.empty:
                # Rename only columns that exist
                cols_to_rename = {k: v for k, v in column_mapping.items() if k in df.columns}
                df.rename(columns=cols_to_rename, inplace=True)
        
        # Ensure all required columns exist
        required_columns = ['Name', 'Status', 'Email address', 'Position', 'Office', 'Building']
        
        for df in [current_df, upcoming_df, past_df]:
            if not df.empty:
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = None
        
        # Set proper Status values
        current_df['Status'] = current_df['Status'].fillna('Current')
        upcoming_df['Status'] = upcoming_df['Status'].fillna('Upcoming')
        past_df['Status'] = past_df['Status'].fillna('Past')
        
        # Convert room numbers to strings for consistency and strip whitespace
        for df in [current_df, upcoming_df, past_df]:
            if not df.empty:
                df['Office'] = df['Office'].astype(str).str.strip()
                if 'Building' in df.columns:
                    df['Building'] = df['Building'].fillna('').astype(str).str.strip()
        
        return current_df, upcoming_df, past_df
    
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# Function to save data to Excel file
def save_data(current_df, upcoming_df, past_df, file_path=DEFAULT_EXCEL_PATH, capacities=None):
    try:
        # Create backup
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f'data/backup/MP_Office_Allocation_{timestamp}.xlsx'
        
        # Copy the original file to backup if it exists
        if os.path.exists(file_path):
            try:
                original_data = pd.read_excel(file_path, sheet_name=None)
                with pd.ExcelWriter(backup_path) as writer:
                    for sheet_name, df in original_data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                st.success(f"Backup created at {backup_path}")
            except Exception as e:
                st.warning(f"Couldn't create backup: {e}")
        
        # Ensure columns are properly formatted before saving
        for df in [current_df, upcoming_df, past_df]:
            if not df.empty:
                # Ensure Building and Office columns exist and are strings
                if 'Building' in df.columns:
                    df['Building'] = df['Building'].astype(str)
                if 'Office' in df.columns:
                    df['Office'] = df['Office'].astype(str)
        
        # Save the modified data
        with pd.ExcelWriter(file_path) as writer:
            current_df.to_excel(writer, sheet_name='Current', index=False)
            upcoming_df.to_excel(writer, sheet_name='Upcoming', index=False)
            past_df.to_excel(writer, sheet_name='Past', index=False)
        
        # Save room capacities if provided
        if capacities is not None:
            save_room_capacities(capacities)
        
        st.success("Data saved successfully!")
        return True
    
    except Exception as e:
        st.error(f"Error saving data: {e}")
        return False


# Load data if not already loaded
if st.session_state.current_df.empty:
    if os.path.exists(st.session_state.file_path):
        st.session_state.current_df, st.session_state.upcoming_df, st.session_state.past_df = load_data(st.session_state.file_path)

# Function to get unique buildings from all dataframes
def get_unique_buildings():
    buildings = set()
    for df in [st.session_state.current_df, st.session_state.upcoming_df, st.session_state.past_df]:
        if not df.empty and 'Building' in df.columns:
            buildings.update(df['Building'].dropna().unique())
    return sorted(list(buildings))

# Function to get unique offices from all dataframes
def get_unique_offices():
    offices = set()
    for df in [st.session_state.current_df, st.session_state.upcoming_df, st.session_state.past_df]:
        if not df.empty and 'Office' in df.columns:
            offices.update(df['Office'].dropna().unique())
    return sorted(list(offices))

# Function to extract floor from room number
def extract_floor(office):
    if isinstance(office, str) and '.' in office:
        try:
            return office.split('.')[0]
        except:
            pass
    return 'Unknown'

# Function to get room occupancy data
def get_room_occupancy_data():
    df = st.session_state.current_df
    
    if df.empty or 'Office' not in df.columns or 'Building' not in df.columns:
        return pd.DataFrame()
    
    # Group by Building and Office to count occupants
    occupancy = df.groupby(['Building', 'Office']).size().reset_index(name='Occupants')
    
    # Add floor information
    occupancy['Floor'] = occupancy['Office'].apply(extract_floor)
    
    # Add storage flag
    occupancy['IsStorage'] = occupancy.apply(
        lambda x: True if any(df[(df['Office'] == x['Office']) & 
                                 (df['Building'] == x['Building'])]['Name'].str.contains('STORAGE', case=False, na=False)) 
                           else False,
        axis=1
    )
    
    # Add capacity information
    occupancy['Max_Capacity'] = occupancy.apply(
        lambda row: st.session_state.room_capacities.get(f"{row['Building']}:{row['Office']}", 2),  # Default to 2
        axis=1
    )
    
    # Calculate capacity metrics
    occupancy['Remaining'] = occupancy['Max_Capacity'] - occupancy['Occupants']
    occupancy['Percentage'] = (occupancy['Occupants'] / occupancy['Max_Capacity'] * 100).round(1)
    
    # Sort by building, floor, and room
    occupancy = occupancy.sort_values(['Building', 'Floor', 'Office'])
    
    return occupancy

# Function to initialize room capacities based on current occupancy
def initialize_room_capacities():
    occupancy_data = get_room_occupancy_data()
    
    if not occupancy_data.empty:
        capacities = {}
        for _, row in occupancy_data.iterrows():
            building = row['Building']
            office = row['Office']
            key = f"{building}:{office}"
            
            # If storage, set capacity to 0
            if row['IsStorage']:
                capacities[key] = 0
            else:
                # Set default max capacity based on current occupancy
                current_occupants = row['Occupants']
                # For rooms with 3+ people, assume that's the capacity
                # For rooms with 0-2 people, set default to 2 unless already occupied by more
                default_capacity = max(current_occupants, 2)
                capacities[key] = default_capacity
        
        return capacities
    
    return {}

# Initialize room capacities if empty
if not st.session_state.room_capacities:
    st.session_state.room_capacities = initialize_room_capacities()
    save_room_capacities(st.session_state.room_capacities)

# Helper function to get room capacity status color class
def get_capacity_class(occupants, max_capacity):
    if max_capacity == 0:  # Storage rooms
        return "room-storage"
    
    if occupants == 0:
        return "room-vacant"
    
    percentage = (occupants / max_capacity * 100) if max_capacity > 0 else 100
    
    if percentage <= 25:
        return "room-low"
    elif percentage <= 50:
        return "room-medium"
    elif percentage <= 75:
        return "room-high"
    else:
        return "room-full"

# Helper function to style status
def style_status(status):
    if isinstance(status, str):
        status_lower = status.lower()
        if 'current' in status_lower:
            return f'<span class="status-current">{status}</span>'
        elif 'upcoming' in status_lower:
            return f'<span class="status-upcoming">{status}</span>'
        elif 'past' in status_lower:
            return f'<span class="status-past">{status}</span>'
    return status

# Sidebar - Settings and Filters
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2329/2329140.png", width=100)
    st.title("Room Allocation System")
    
    # File Upload/Selection
    st.header("Data Source")
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        # Save the uploaded file
        with open('data/temp_upload.xlsx', 'wb') as f:
            f.write(uploaded_file.getvalue())
        
        st.session_state.file_path = 'data/temp_upload.xlsx'
        st.session_state.current_df, st.session_state.upcoming_df, st.session_state.past_df = load_data(st.session_state.file_path)
        
        # Initialize room capacities if this is a new file
        if not st.session_state.room_capacities:
            st.session_state.room_capacities = initialize_room_capacities()
            save_room_capacities(st.session_state.room_capacities)
            
        st.success("File uploaded successfully!")
    
    # Filters
    st.header("Filters")
    buildings = ['All'] + get_unique_buildings()
    st.session_state.filter_building = st.selectbox("Building", buildings)
    
    # Navigation
    st.header("Navigation")
    page = st.radio("Go to", ["Dashboard", "Current Occupants", "Upcoming Occupants" 
                             , "Room Management", "Reports"])
    
    # Save button with improved data validation
    st.header("Actions")
    if st.button("💾 Save Changes"):
        # Add data validation before saving
        validation_ok = True
        validation_messages = []
        
        # Check for required fields in current occupants
        if not st.session_state.current_df.empty:
            # Check for missing building or office values
            missing_location = st.session_state.current_df[
                (st.session_state.current_df['Building'].isna()) | 
                (st.session_state.current_df['Building'] == "") |
                (st.session_state.current_df['Office'].isna()) | 
                (st.session_state.current_df['Office'] == "")
            ]
            
            if not missing_location.empty and len(missing_location) > 0:
                validation_ok = False
                validation_messages.append(f"⚠️ {len(missing_location)} current occupants are missing Building or Office assignment.")
        
        # Proceed with save if validation passes
        if validation_ok or st.checkbox("Save anyway (ignore warnings)", key="ignore_warnings"):
            success = save_data(
                st.session_state.current_df, 
                st.session_state.upcoming_df, 
                st.session_state.past_df,
                st.session_state.file_path,
                st.session_state.room_capacities
            )
            if success:
                st.session_state.last_save = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        else:
            for msg in validation_messages:
                st.warning(msg)
            st.error("Please fix the issues above or check 'Save anyway' to proceed.")
    
    # Add data backup button
    if st.button("📦 Create Backup"):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f'data/backup/MP_Office_Allocation_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(backup_path) as writer:
                st.session_state.current_df.to_excel(writer, sheet_name='Current', index=False)
                st.session_state.upcoming_df.to_excel(writer, sheet_name='Upcoming', index=False)
                st.session_state.past_df.to_excel(writer, sheet_name='Past', index=False)
            
            st.success(f"Backup created: {backup_path}")
        except Exception as e:
            st.error(f"Error creating backup: {e}")
    
    # Display last save time
    if st.session_state.last_save:
        st.info(f"Last saved: {st.session_state.last_save}")
    
    st.markdown("---")
    st.caption("Office Room Allocation System v2.1")

# Dashboard Page
if page == "Dashboard":
    st.title("Office Allocation Dashboard")
    
    # Summary Statistics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        current_count = len(st.session_state.current_df)
        st.metric("Current Occupants", current_count)
    
    with col2:
        upcoming_count = len(st.session_state.upcoming_df)
        st.metric("Upcoming Occupants", upcoming_count)
    
    with col3:
        past_count = len(st.session_state.past_df)
        st.metric("Past Occupants", past_count)
    
    with col4:
        # Calculate capacity metrics
        room_occupancy = get_room_occupancy_data()
        if not room_occupancy.empty:
            total_capacity = room_occupancy['Max_Capacity'].sum()
            current_occupancy = room_occupancy['Occupants'].sum()
            occupancy_percentage = (current_occupancy / total_capacity * 100) if total_capacity > 0 else 0
            st.metric("Occupancy Rate", f"{occupancy_percentage:.1f}%")
        else:
            st.metric("Occupancy Rate", "0.0%")
    
    # Main dashboard sections
    col_left, col_right = st.columns([3, 2])
    
    with col_left:
        st.subheader("Room Occupancy by Building and Floor")
        
        # Get room occupancy data
        room_occupancy = get_room_occupancy_data()
        
        if not room_occupancy.empty:
            # Filter by building if needed
            if st.session_state.filter_building != 'All':
                room_occupancy = room_occupancy[room_occupancy['Building'] == st.session_state.filter_building]
            
            # Group by building and floor for the visualization
            building_floor_occupancy = room_occupancy.groupby(['Building', 'Floor']).agg({
                'Occupants': 'sum',
                'Max_Capacity': 'sum',
                'Office': 'count'
            }).reset_index()
            
            building_floor_occupancy.rename(columns={'Office': 'Room Count'}, inplace=True)
            building_floor_occupancy['Occupancy Rate'] = (building_floor_occupancy['Occupants'] / 
                                                         building_floor_occupancy['Max_Capacity'] * 100).round(1)
            
            # Show the summary table
            st.dataframe(building_floor_occupancy, use_container_width=True)
            
            # Create a detailed occupancy visualization
            fig = go.Figure()
            
            for building in building_floor_occupancy['Building'].unique():
                building_data = building_floor_occupancy[building_floor_occupancy['Building'] == building]
                
                fig.add_trace(go.Bar(
                    x=building_data['Floor'],
                    y=building_data['Occupants'],
                    name=f"{building} - Occupants",
                    marker_color='#4CAF50',
                    text=building_data['Occupants'],
                    textposition='auto',
                    hovertemplate=
                    '<b>%{x}</b><br>' +
                    'Building: ' + building + '<br>' +
                    'Occupants: %{y}<br>' +
                    'Capacity: %{customdata[0]}<br>' +
                    'Occupancy Rate: %{customdata[1]}%<br>' +
                    'Room Count: %{customdata[2]}',
                    customdata=np.stack((
                        building_data['Max_Capacity'], 
                        building_data['Occupancy Rate'],
                        building_data['Room Count']
                    ), axis=1)
                ))
                
                fig.add_trace(go.Bar(
                    x=building_data['Floor'],
                    y=building_data['Max_Capacity'] - building_data['Occupants'],
                    name=f"{building} - Available",
                    marker_color='#FFC107',
                    text=building_data['Max_Capacity'] - building_data['Occupants'],
                    textposition='auto',
                    hoverinfo='skip'
                ))
            
            fig.update_layout(
                barmode='stack',
                title='Occupancy by Building and Floor',
                xaxis_title='Floor',
                yaxis_title='Number of People',
                height=500,
                legend_title='Occupancy Status',
                hovermode='closest'
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Room capacity distribution
            st.subheader("Room Capacity Utilization")
            
            # Create capacity categories
            room_occupancy['Capacity_Category'] = room_occupancy.apply(
                lambda row: 'Vacant' if row['Occupants'] == 0 
                            else 'Low (1-25%)' if row['Percentage'] <= 25
                            else 'Medium (26-50%)' if row['Percentage'] <= 50
                            else 'High (51-75%)' if row['Percentage'] <= 75
                            else 'Very High (76-99%)' if row['Percentage'] < 100
                            else 'Full (100%)',
                axis=1
            )
            
            # Create color mapping
            color_map = {
                'Vacant': '#d4edda',
                'Low (1-25%)': '#e6f7e1',
                'Medium (26-50%)': '#fff3cd',
                'High (51-75%)': '#ffe5d9',
                'Very High (76-99%)': '#ffcccc',
                'Full (100%)': '#f8d7da'
            }
            
            # Group by capacity category
            capacity_counts = room_occupancy['Capacity_Category'].value_counts().reset_index()
            capacity_counts.columns = ['Capacity', 'Count']
            
            # Order categories
            category_order = ['Vacant', 'Low (1-25%)', 'Medium (26-50%)', 'High (51-75%)', 'Very High (76-99%)', 'Full (100%)']
            capacity_counts['Capacity'] = pd.Categorical(
                capacity_counts['Capacity'], 
                categories=category_order, 
                ordered=True
            )
            capacity_counts = capacity_counts.sort_values('Capacity')
            
            fig = px.bar(
                capacity_counts,
                x='Capacity',
                y='Count',
                color='Capacity',
                color_discrete_map=color_map,
                text='Count',
                title='Room Capacity Utilization'
            )
            
            fig.update_layout(
                xaxis_title='Capacity Utilization',
                yaxis_title='Number of Rooms',
                height=400,
                showlegend=False
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No occupancy data available")
    
    with col_right:
        # Status Distribution
        st.subheader("Occupant Status Distribution")
        
        if not st.session_state.current_df.empty and 'Status' in st.session_state.current_df.columns:
            status_counts = st.session_state.current_df['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            
            # Consolidate rare statuses
            if len(status_counts) > 5:
                main_statuses = status_counts.head(4)
                other_count = status_counts.tail(len(status_counts) - 4)['Count'].sum()
                main_statuses = pd.concat([
                    main_statuses, 
                    pd.DataFrame([{'Status': 'Other', 'Count': other_count}])
                ])
                status_counts = main_statuses
            
            fig = px.pie(
                status_counts,
                values='Count',
                names='Status',
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Safe
            )
            
            fig.update_layout(height=300)
            fig.update_traces(textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No status data available")
        
        # Buildings Distribution
        st.subheader("Occupants by Building")
        
        if not st.session_state.current_df.empty and 'Building' in st.session_state.current_df.columns:
            building_counts = st.session_state.current_df['Building'].value_counts().reset_index()
            building_counts.columns = ['Building', 'Count']
            
            fig = px.bar(
                building_counts,
                x='Building',
                y='Count',
                color='Building',
                text='Count',
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            
            fig.update_layout(
                xaxis_title='Building Name',
                yaxis_title='Number of Occupants',
                height=300,
                showlegend=False
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No building data available")
        
        # Position/Role Distribution
        st.subheader("Occupants by Position")
        
        if not st.session_state.current_df.empty and 'Position' in st.session_state.current_df.columns:
            # Count positions, handle NaN values
            position_counts = st.session_state.current_df['Position'].fillna('Unknown').value_counts()
            
            # Only keep the top 6 positions, group others
            if len(position_counts) > 6:
                top_positions = position_counts.head(5)
                other_count = position_counts.tail(len(position_counts) - 5).sum()
                position_counts = pd.concat([top_positions, pd.Series([other_count], index=['Other'])])
            
            position_df = position_counts.reset_index()
            position_df.columns = ['Position', 'Count']
            
            fig = px.bar(
                position_df,
                x='Count',
                y='Position',
                color='Count',
                orientation='h',
                text='Count',
                color_continuous_scale='Viridis'
            )
            
            fig.update_layout(
                xaxis_title='Number of Occupants',
                yaxis_title='Position',
                height=300,
                yaxis={'categoryorder': 'total ascending'},
                coloraxis_showscale=False
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No position data available")
    
    # Quick access to upcoming occupants
    if not st.session_state.upcoming_df.empty:
        st.subheader("Upcoming Occupants")
        upcoming_preview = st.session_state.upcoming_df.head(5)
        st.dataframe(upcoming_preview, use_container_width=True)
        
        if len(st.session_state.upcoming_df) > 5:
            st.caption(f"Showing 5 of {len(st.session_state.upcoming_df)} upcoming occupants. Go to Upcoming Occupants page to see all.")

# Current Occupants Page
elif page == "Current Occupants":
    st.title("Current Occupants")
    
    # Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        filter_name = st.text_input("Filter by Name")
    with col2:
        if st.session_state.filter_building != 'All':
            building_filter = st.session_state.filter_building
        else:
            building_filter = st.selectbox("Filter by Building", ['All'] + get_unique_buildings())
    with col3:
        office_filter = st.selectbox("Filter by Office", ['All'] + get_unique_offices())
    
    # Apply filters
    filtered_df = st.session_state.current_df.copy()
    
    if filter_name:
        filtered_df = filtered_df[filtered_df['Name'].str.contains(filter_name, case=False, na=False)]
    
    if building_filter != 'All':
        filtered_df = filtered_df[filtered_df['Building'] == building_filter]
    
    if office_filter != 'All':
        filtered_df = filtered_df[filtered_df['Office'] == office_filter]
    
    # Display data and allow editing
    if not filtered_df.empty:
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Current", "Upcoming", "Past"],
                    required=True
                ),
                "Building": st.column_config.SelectboxColumn(
                    "Building",
                    options=get_unique_buildings(),
                    required=True
                ),
                "Email address": st.column_config.TextColumn(
                    "Email address",
                    help="User's email address"
                ),
                "Office": st.column_config.SelectboxColumn(
                    "Office",
                    options=get_unique_offices(),
                    required=True
                ),
                "Position": st.column_config.TextColumn(
                    "Position",
                    help="Occupant's role or position"
                )
            },
            hide_index=True,
            num_rows="dynamic"
        )
        
        # Update the dataframe if changes were made
        if not edited_df.equals(filtered_df):
            # Handle status changes that require moving between sheets
            for i, row in edited_df.iterrows():
                orig_idx = filtered_df.index[filtered_df.reset_index().index == i][0]
                
                # Check if status has changed
                if 'Status' in row and row['Status'] != filtered_df.loc[orig_idx, 'Status']:
                    # If status changed to Upcoming
                    if row['Status'] == 'Upcoming':
                        # Add to upcoming_df
                        st.session_state.upcoming_df = pd.concat([st.session_state.upcoming_df, pd.DataFrame([row])], ignore_index=True)
                        # Remove from current_df
                        st.session_state.current_df = st.session_state.current_df.drop(orig_idx)
                    # If status changed to Past
                    elif row['Status'] == 'Past':
                        # Add to past_df
                        st.session_state.past_df = pd.concat([st.session_state.past_df, pd.DataFrame([row])], ignore_index=True)
                        # Remove from current_df
                        st.session_state.current_df = st.session_state.current_df.drop(orig_idx)
                else:
                    # Just update the row
                    st.session_state.current_df.loc[orig_idx] = row
            
            st.success("Changes applied! Remember to click 'Save Changes' in the sidebar to save them permanently.")
    else:
        st.info("No current occupants found with the selected filters")
    
    # Add new occupant section
    st.markdown("---")
    st.subheader("Add New Occupant")
    
    with st.form("add_occupant_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            new_name = st.text_input("Name", placeholder="e.g., Smith, John Dr")
            new_email = st.text_input("Email", placeholder="e.g., john.smith@anu.edu.au")
            new_position = st.text_input("Position", placeholder="e.g., Professor")
        
        with col2:
            # Get room occupancy for intelligent room assignment
            room_occupancy = get_room_occupancy_data()
            
            if not room_occupancy.empty:
                # Add formatted room labels with capacity info
                room_options = []
                
                for _, room in room_occupancy.iterrows():
                    building = room['Building']
                    office = room['Office']
                    occupants = room['Occupants']
                    max_capacity = room['Max_Capacity']
                    remaining = room['Remaining']
                    percentage = room['Percentage']
                    
                    # Skip storage rooms
                    if room['IsStorage']:
                        continue
                    
                    # Create formatted room label
                    status_label = (f"Vacant" if occupants == 0 else 
                                    f"{occupants}/{max_capacity} occupants ({percentage}%)")
                    
                    room_label = f"{building} - {office} [{status_label}]"
                    room_options.append((room_label, building, office, percentage))
                
                # Sort by occupancy percentage (lowest first)
                room_options.sort(key=lambda x: x[3])
                
                # Create dropdown with room options
                if room_options:
                    selected_room = st.selectbox(
                        "Room (sorted by availability)",
                        [option[0] for option in room_options]
                    )
                    
                    # Extract building and office from selection
                    selected_index = [option[0] for option in room_options].index(selected_room)
                    new_building = room_options[selected_index][1]
                    new_office = room_options[selected_index][2]
                    
                    # Show capacity warning if needed
                    selected_percentage = room_options[selected_index][3]
                    if selected_percentage >= 75:
                        st.warning(f"This room is at {selected_percentage}% capacity")
                    elif selected_percentage >= 100:
                        st.error("This room is at full capacity")
                else:
                    new_office = st.text_input("Office", placeholder="e.g., 3.17")
                    new_building = st.selectbox("Building", get_unique_buildings())
            else:
                new_office = st.text_input("Office", placeholder="e.g., 3.17")
                new_building = st.selectbox("Building", get_unique_buildings())
            
            new_status = st.selectbox("Status", ["Current", "Upcoming", "Past"], index=0)
        
        submitted = st.form_submit_button("Add Occupant")
        
        if submitted:
            if new_name and new_office and new_building:
                # Create new row
                new_row = {
                    'Name': new_name,
                    'Email address': new_email,
                    'Position': new_position,
                    'Office': new_office,
                    'Building': new_building,
                    'Status': new_status
                }
                
                # Add to appropriate dataframe
                if new_status == "Current":
                    st.session_state.current_df = pd.concat([st.session_state.current_df, pd.DataFrame([new_row])], ignore_index=True)
                elif new_status == "Upcoming":
                    st.session_state.upcoming_df = pd.concat([st.session_state.upcoming_df, pd.DataFrame([new_row])], ignore_index=True)
                elif new_status == "Past":
                    st.session_state.past_df = pd.concat([st.session_state.past_df, pd.DataFrame([new_row])], ignore_index=True)
                
                st.success(f"Added {new_name} to {new_status} occupants. Remember to save changes!")
            else:
                st.error("Name, Office, and Building are required fields")

# Upcoming Occupants Page
elif page == "Upcoming Occupants":
    st.title("Upcoming Occupants")
    
    # Filters
    col1, col2 = st.columns(2)
    with col1:
        filter_name = st.text_input("Filter by Name")
    with col2:
        building_filter = st.selectbox("Filter by Building", ['All'] + get_unique_buildings())
    
    # Apply filters
    filtered_df = st.session_state.upcoming_df.copy()
    
    if filter_name:
        filtered_df = filtered_df[filtered_df['Name'].str.contains(filter_name, case=False, na=False)]
    
    if building_filter != 'All':
        if 'Building' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Building'] == building_filter]
    
    # Display data and allow editing
    if not filtered_df.empty:
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Current", "Upcoming", "Past"],
                    required=True
                ),
                "Building": st.column_config.SelectboxColumn(
                    "Building",
                    options=get_unique_buildings(),
                    required=True
                ),
                "Office": st.column_config.SelectboxColumn(
                    "Office",
                    options=get_unique_offices(),
                    required=True
                )
            },
            hide_index=True,
            num_rows="dynamic"
        )
        
        # Update the dataframe if changes were made
        if not edited_df.equals(filtered_df):
            # Handle moving between sheets based on status
            for i, row in edited_df.iterrows():
                # Find the original row index
                orig_idx = filtered_df.index[filtered_df.reset_index().index == i][0]
                
                # Check if status has changed
                if 'Status' in row and row['Status'] != filtered_df.loc[orig_idx, 'Status']:
                    # If status changed to Current
                    if row['Status'] == 'Current':
                        # Add to current_df
                        st.session_state.current_df = pd.concat([st.session_state.current_df, pd.DataFrame([row])], ignore_index=True)
                        # Remove from upcoming_df
                        st.session_state.upcoming_df = st.session_state.upcoming_df.drop(orig_idx)
                    # If status changed to Past
                    elif row['Status'] == 'Past':
                        # Add to past_df
                        st.session_state.past_df = pd.concat([st.session_state.past_df, pd.DataFrame([row])], ignore_index=True)
                        # Remove from upcoming_df
                        st.session_state.upcoming_df = st.session_state.upcoming_df.drop(orig_idx)
                else:
                    # Just update the row
                    st.session_state.upcoming_df.loc[orig_idx] = row
            
            st.success("Changes applied! Remember to click 'Save Changes' in the sidebar to save them permanently.")
    else:
        st.info("No upcoming occupants found with the selected filters")
    
    # Add new upcoming occupant section
    st.markdown("---")
    st.subheader("Add New Upcoming Occupant")
    
    with st.form("add_upcoming_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            new_name = st.text_input("Name", placeholder="e.g., Smith, John Dr")
            new_email = st.text_input("Email", placeholder="e.g., john.smith@anu.edu.au")
            new_position = st.text_input("Position", placeholder="e.g., Professor")
            planned_arrival = st.date_input("Planned Arrival Date")
        
        with col2:
            # Get room occupancy data for intelligent room suggestion
            room_occupancy = get_room_occupancy_data()
            
            if not room_occupancy.empty:
                # Add rooms with availability info
                room_options = []
                
                for _, room in room_occupancy.iterrows():
                    building = room['Building']
                    office = room['Office']
                    occupants = room['Occupants']
                    max_capacity = room['Max_Capacity']
                    remaining = room['Remaining']
                    
                    # Skip storage rooms
                    if room['IsStorage']:
                        continue
                    
                    # Skip full rooms
                    if remaining <= 0:
                        continue
                    
                    # Create formatted room label
                    status_label = f"{occupants}/{max_capacity} occupants ({remaining} available)"
                    room_label = f"{building} - {office} [{status_label}]"
                    
                    room_options.append((room_label, building, office, remaining))
                
                # Sort by most available space
                room_options.sort(key=lambda x: x[3], reverse=True)
                
                # Create dropdown with room options
                if room_options:
                    selected_room = st.selectbox(
                        "Recommended Room (sorted by availability)",
                        [option[0] for option in room_options]
                    )
                    
                    # Extract building and office from selection
                    selected_index = [option[0] for option in room_options].index(selected_room)
                    new_building = room_options[selected_index][1]
                    new_office = room_options[selected_index][2]
                else:
                    new_office = st.text_input("Office", placeholder="e.g., 3.17")
                    new_building = st.selectbox("Building", get_unique_buildings())
            else:
                new_office = st.text_input("Office", placeholder="e.g., 3.17") 
                new_building = st.selectbox("Building", get_unique_buildings())
        
        submitted = st.form_submit_button("Add Upcoming Occupant")
        
        if submitted:
            if new_name:
                # Create new row
                new_row = {
                    'Name': new_name,
                    'Email address': new_email,
                    'Position': new_position,
                    'Office': new_office,
                    'Building': new_building,
                    'Status': 'Upcoming',
                    'Planned Arrival': planned_arrival
                }
                
                # Add to upcoming dataframe
                st.session_state.upcoming_df = pd.concat([st.session_state.upcoming_df, pd.DataFrame([new_row])], ignore_index=True)
                st.success(f"Added {new_name} to upcoming occupants. Remember to save changes!")
            else:
                st.error("Name is required")

# Past Occupants Page
elif page == "Past Occupants":
    st.title("Past Occupants")
    
    # Filters
    col1, col2 = st.columns(2)
    with col1:
        filter_name = st.text_input("Filter by Name")
    with col2:
        building_filter = st.selectbox("Filter by Building", ['All'] + get_unique_buildings())
    
    # Apply filters
    filtered_df = st.session_state.past_df.copy()
    
    if filter_name:
        filtered_df = filtered_df[filtered_df['Name'].str.contains(filter_name, case=False, na=False)]
    
    if building_filter != 'All' and 'Building' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Building'] == building_filter]
    
    # Display data
    if not filtered_df.empty:
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Current", "Upcoming", "Past"],
                    required=True
                )
            },
            hide_index=True
        )
        
        # Update the dataframe if changes were made
        if not edited_df.equals(filtered_df):
            # Handle moving between sheets based on status
            for i, row in edited_df.iterrows():
                # Find the original row index
                orig_idx = filtered_df.index[filtered_df.reset_index().index == i][0]
                
                # Check if status has changed
                if 'Status' in row and row['Status'] != filtered_df.loc[orig_idx, 'Status']:
                    # If status changed to Current
                    if row['Status'] == 'Current':
                        # Add to current_df
                        st.session_state.current_df = pd.concat([st.session_state.current_df, pd.DataFrame([row])], ignore_index=True)
                        # Remove from past_df
                        st.session_state.past_df = st.session_state.past_df.drop(orig_idx)
                    # If status changed to Upcoming
                    elif row['Status'] == 'Upcoming':
                        # Add to upcoming_df
                        st.session_state.upcoming_df = pd.concat([st.session_state.upcoming_df, pd.DataFrame([row])], ignore_index=True)
                        # Remove from past_df
                        st.session_state.past_df = st.session_state.past_df.drop(orig_idx)
                else:
                    # Just update the row
                    st.session_state.past_df.loc[orig_idx] = row
            
            st.success("Changes applied! Remember to click 'Save Changes' in the sidebar to save them permanently.")
    else:
        st.info("No past occupants found with the selected filters")
    
    # Add search functionality for past occupants
    st.markdown("---")
    st.subheader("Search Past Occupants")
    
    search_term = st.text_input("Search by name, position, or email", key="past_search")
    
    if search_term:
        search_results = st.session_state.past_df[
            st.session_state.past_df['Name'].str.contains(search_term, case=False, na=False) |
            st.session_state.past_df['Position'].str.contains(search_term, case=False, na=False) |
            st.session_state.past_df['Email address'].str.contains(search_term, case=False, na=False)
        ]
        
        if not search_results.empty:
            st.dataframe(search_results, use_container_width=True)
        else:
            st.info(f"No past occupants found matching '{search_term}'")

# Room Management Page

# Room Management Page
elif page == "Room Management":
    st.title("Room Management")
    
    # Create tabs for different management views
    tab1, tab2, tab3, tab4 = st.tabs(["Room Occupancy", "Edit Rooms", "Room Status", "Room Assignment"])
    
    # Tab 1: Room Occupancy
    with tab1:
        st.subheader("Room Occupancy Overview")
        
        # Filter by building
        building_filter = st.selectbox("Select Building", ['All'] + get_unique_buildings(), key='room_building_filter')
        
        # Get room occupancy data
        room_occupancy = get_room_occupancy_data()
        
        if not room_occupancy.empty:
            # Filter by building if needed
            if building_filter != 'All':
                room_occupancy = room_occupancy[room_occupancy['Building'] == building_filter]
            
            # Show occupancy data with capacity information
            st.write("**Room Occupancy Table**")
            
            # Create a formatted dataframe for display
            display_df = room_occupancy[['Building', 'Floor', 'Office', 'Occupants', 
                                         'Max_Capacity', 'Remaining', 'Percentage']]
            
            # Sort by building, floor, and room number
            display_df = display_df.sort_values(['Building', 'Floor', 'Office'])
            
            # Function to color cells based on occupancy percentage
            def color_occupancy_percentage(val):
                """
                Color cells based on occupancy percentage.
                """
                try:
                    percentage = float(val)
                    if percentage == 0:
                        return 'background-color: #d4edda'  # Vacant - green
                    elif percentage <= 25:
                        return 'background-color: #e6f7e1'  # Low - light green
                    elif percentage <= 50:
                        return 'background-color: #fff3cd'  # Medium - yellow
                    elif percentage <= 75:
                        return 'background-color: #ffe5d9'  # High - light orange
                    elif percentage < 100:
                        return 'background-color: #ffcccc'  # Very high - pink
                    else:
                        return 'background-color: #f8d7da'  # Full - red
                except (ValueError, TypeError):
                    return ''

            # Create display dataframe
            display_df = room_occupancy[['Building', 'Floor', 'Office', 'Occupants', 
                                        'Max_Capacity', 'Remaining', 'Percentage']]

            # Use a more direct styling approach
            styled_df = display_df.style.applymap(
                lambda x: color_occupancy_percentage(x) if isinstance(x, (int, float)) else ''
            )

            # Format percentage column
            styled_df = styled_df.format({'Percentage': '{:.1f}%'})
            styled_df = styled_df.set_table_attributes('class="dataframe"')

            st.dataframe(styled_df, use_container_width=True)
            
            # Show rooms organized by floor
            st.write("**Rooms by Floor**")
            
            for building in room_occupancy['Building'].unique():
                if building_filter != 'All' and building != building_filter:
                    continue
                
                st.markdown(f"### {building}")
                
                building_rooms = room_occupancy[room_occupancy['Building'] == building]
                
                # Get all floors in this building
                floors = sorted(building_rooms['Floor'].unique(), 
                               key=lambda x: float(x) if x.isdigit() or x.replace('.', '', 1).isdigit() else float('inf'))
                
                for floor in floors:
                    st.markdown(f"<div class='floor-heading'>Floor {floor}</div>", unsafe_allow_html=True)
                    
                    # Get rooms on this floor
                    floor_rooms = building_rooms[building_rooms['Floor'] == floor]
                    
                    # Create a grid of rooms
                    cols = st.columns(4)  # 4 rooms per row
                    
                    for i, (_, room) in enumerate(floor_rooms.iterrows()):
                        col_idx = i % 4
                        
                        # Get occupancy information
                        office = room['Office']
                        occupants = room['Occupants']
                        max_capacity = room['Max_Capacity']
                        percentage = room['Percentage']
                        is_storage = room['IsStorage']
                        
                        # Determine status class
                        if is_storage:
                            status_class = "room-storage"
                            status_text = "Storage"
                        elif occupants == 0:
                            status_class = "room-vacant"
                            status_text = f"Vacant (0/{max_capacity})"
                        elif percentage <= 25:
                            status_class = "room-low"
                            status_text = f"Low ({occupants}/{max_capacity})"
                        elif percentage <= 50:
                            status_class = "room-medium"
                            status_text = f"Medium ({occupants}/{max_capacity})"
                        elif percentage <= 75:
                            status_class = "room-high"
                            status_text = f"High ({occupants}/{max_capacity})"
                        elif percentage < 100:
                            status_class = "room-high"
                            status_text = f"Very High ({occupants}/{max_capacity})"
                        else:
                            status_class = "room-full"
                            status_text = f"Full ({occupants}/{max_capacity})"
                        
                        # Create room card
                        with cols[col_idx]:
                            st.markdown(
                                f"""
                                <div class="{status_class}">
                                    <h4 style="margin:0;">{office}</h4>
                                    <p style="margin:0;">{status_text}</p>
                                    <p style="margin:0; font-size:0.8em;">Capacity: {percentage:.1f}%</p>
                                </div>
                                """, 
                                unsafe_allow_html=True
                            )
                            
                            # Show occupants if any
                            if occupants > 0 and not is_storage:
                                room_occupants = st.session_state.current_df[
                                    (st.session_state.current_df['Building'] == room['Building']) & 
                                    (st.session_state.current_df['Office'] == room['Office'])
                                ]
                                
                                for _, occupant in room_occupants.iterrows():
                                    st.markdown(
                                        f"<span class='occupant-tag'>{occupant['Name']}</span>",
                                        unsafe_allow_html=True
                                    )
        else:
            st.info("No room data available")
    
    # Tab 2: Fully Editable Room Table
    with tab2:
        st.subheader("Edit Room Information")
        
        # Get room occupancy data
        room_occupancy = get_room_occupancy_data()
        
        if not room_occupancy.empty:
            # Create a more comprehensive dataframe for editing that includes all needed fields
            edit_df = room_occupancy.copy()
            
            # Add room type column
            edit_df['Room Type'] = edit_df['IsStorage'].apply(lambda x: 'Storage' if x else 'Regular')
            
            # Create a fully editable version with better column names and ordering
            editable_df = edit_df[[
                'Building', 'Office', 'Floor', 'Occupants', 'Max_Capacity', 
                'Remaining', 'Percentage', 'Room Type'
            ]].copy()
            
            # Rename columns for clarity
            editable_df = editable_df.rename(columns={
                'Max_Capacity': 'Capacity',
                'Percentage': 'Occupancy %'
            })
            
            # Allow building filtering for easier management
            building_filter = st.selectbox(
                "Filter by Building", 
                ['All'] + sorted(editable_df['Building'].unique().tolist()),
                key='edit_building_filter'
            )
            
            if building_filter != 'All':
                filtered_df = editable_df[editable_df['Building'] == building_filter].copy()
            else:
                filtered_df = editable_df.copy()
            
            # Sort for easier viewing
            filtered_df = filtered_df.sort_values(['Building', 'Floor', 'Office'])
            
            # Show current data in an editable table
            st.write("Edit the table directly by clicking on cells. All fields are editable except Occupants, Remaining, and Occupancy %.")
            st.write("To add a new room, add a new row. To delete a room, remove all text from that row.")
            
            edited_df = st.data_editor(
                filtered_df,
                column_config={
                    "Building": st.column_config.TextColumn(
                        "Building",
                        help="Building name - edit directly to rename"
                    ),
                    "Office": st.column_config.TextColumn(
                        "Office",
                        help="Room number"
                    ),
                    "Floor": st.column_config.TextColumn(
                        "Floor",
                        help="Floor number (derived from room number)"
                    ),
                    "Capacity": st.column_config.NumberColumn(
                        "Capacity",
                        min_value=0,
                        max_value=20,
                        help="Maximum number of people allowed in this room"
                    ),
                    "Occupants": st.column_config.NumberColumn(
                        "Occupants",
                        disabled=True,
                        help="Current number of occupants (read-only)"
                    ),
                    "Remaining": st.column_config.NumberColumn(
                        "Remaining",
                        disabled=True,
                        help="Available space in the room (read-only)"
                    ),
                    "Occupancy %": st.column_config.ProgressColumn(
                        "Occupancy %",
                        format="%.1f%%",
                        min_value=0,
                        max_value=100,
                        help="Percentage of capacity used (read-only)"
                    ),
                    "Room Type": st.column_config.SelectboxColumn(
                        "Room Type",
                        options=["Regular", "Storage"],
                        help="Type of room"
                    )
                },
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="room_editor"
            )
            
            # Add a save button for the edited data
            if st.button("Save Room Changes", type="primary"):
                try:
                    # Track changes to process
                    changes_made = False
                    changes_summary = []
                    
                    # Get original keys from room_capacities for comparison
                    original_keys = set(st.session_state.room_capacities.keys())
                    new_keys = set()
                    
                    # Process edited rows
                    for i, row in edited_df.iterrows():
                        # Skip empty rows (these are considered deleted)
                        if pd.isna(row['Building']) or pd.isna(row['Office']) or row['Building'] == '' or row['Office'] == '':
                            continue
                            
                        building = str(row['Building']).strip()
                        office = str(row['Office']).strip()
                        capacity = int(row['Capacity']) if not pd.isna(row['Capacity']) else 0
                        room_type = row['Room Type']
                        is_storage = room_type == 'Storage'
                        
                        # Generate the room key
                        room_key = f"{building}:{office}"
                        new_keys.add(room_key)
                        
                        # Check if this is a new room or updated room
                        is_new = room_key not in original_keys
                        
                        # Update capacity in all cases
                        st.session_state.room_capacities[room_key] = capacity
                        
                        # For new rooms, we need to add a placeholder record
                        if is_new:
                            changes_made = True
                            changes_summary.append(f"Added new room: {building} - {office}")
                            
                            # Add appropriate placeholder depending on room type
                            if is_storage:
                                new_record = {
                                    'Name': 'STORAGE',
                                    'Status': 'Current',
                                    'Email address': '',
                                    'Position': '',
                                    'Office': office,
                                    'Building': building
                                }
                            else:
                                new_record = {
                                    'Name': 'PLACEHOLDER',
                                    'Status': 'Current',
                                    'Email address': '',
                                    'Position': '',
                                    'Office': office,
                                    'Building': building
                                }
                            
                            # Add to current occupants
                            st.session_state.current_df = pd.concat(
                                [st.session_state.current_df, pd.DataFrame([new_record])], 
                                ignore_index=True
                            )
                        else:
                            # Check if room type has changed
                            original_row = edit_df[(edit_df['Building'] == building) & (edit_df['Office'] == office)]
                            
                            if not original_row.empty:
                                original_is_storage = original_row.iloc[0]['IsStorage']
                                
                                if original_is_storage != is_storage:
                                    changes_made = True
                                    changes_summary.append(f"Changed room type for {building} - {office}")
                                    
                                    # Handle the changes in room type
                                    if is_storage:
                                        # Convert to storage - add STORAGE record if not exists
                                        storage_exists = False
                                        if not st.session_state.current_df.empty:
                                            storage_mask = (
                                                (st.session_state.current_df['Building'] == building) & 
                                                (st.session_state.current_df['Office'] == office) & 
                                                (st.session_state.current_df['Name'].fillna('').str.contains('STORAGE', case=False))
                                            )
                                            storage_exists = storage_mask.any()
                                        
                                        if not storage_exists:
                                            # Add a STORAGE record
                                            storage_record = {
                                                'Name': 'STORAGE',
                                                'Status': 'Current',
                                                'Email address': '',
                                                'Position': '',
                                                'Office': office,
                                                'Building': building
                                            }
                                            st.session_state.current_df = pd.concat(
                                                [st.session_state.current_df, pd.DataFrame([storage_record])], 
                                                ignore_index=True
                                            )
                                    else:
                                        # Convert from storage to regular - remove STORAGE records
                                        if not st.session_state.current_df.empty:
                                            storage_mask = (
                                                (st.session_state.current_df['Building'] == building) & 
                                                (st.session_state.current_df['Office'] == office) & 
                                                (st.session_state.current_df['Name'].fillna('').str.contains('STORAGE', case=False))
                                            )
                                            
                                            if storage_mask.any():
                                                st.session_state.current_df = st.session_state.current_df[~storage_mask].copy()
                                                
                                                # Add a placeholder if no other occupants
                                                remaining_occupants = st.session_state.current_df[
                                                    (st.session_state.current_df['Building'] == building) & 
                                                    (st.session_state.current_df['Office'] == office)
                                                ]
                                                
                                                if remaining_occupants.empty:
                                                    placeholder = {
                                                        'Name': 'PLACEHOLDER',
                                                        'Status': 'Current',
                                                        'Email address': '',
                                                        'Position': '',
                                                        'Office': office,
                                                        'Building': building
                                                    }
                                                    st.session_state.current_df = pd.concat(
                                                        [st.session_state.current_df, pd.DataFrame([placeholder])], 
                                                        ignore_index=True
                                                    )
                    
                    # Find deleted rooms by comparing original keys to new keys
                    deleted_keys = original_keys - new_keys
                    
                    for key in deleted_keys:
                        changes_made = True
                        building, office = key.split(':')
                        changes_summary.append(f"Deleted room: {building} - {office}")
                        
                        # Remove from room capacities
                        if key in st.session_state.room_capacities:
                            del st.session_state.room_capacities[key]
                        
                        # Remove from all dataframes
                        dataframes = [
                            st.session_state.current_df, 
                            st.session_state.upcoming_df, 
                            st.session_state.past_df
                        ]
                        
                        for i, df in enumerate(dataframes):
                            if not df.empty:
                                room_mask = (df['Building'] == building) & (df['Office'] == office)
                                if room_mask.any():
                                    # Store the updated dataframe back to session state
                                    if i == 0:
                                        st.session_state.current_df = df[~room_mask].copy()
                                    elif i == 1:
                                        st.session_state.upcoming_df = df[~room_mask].copy()
                                    else:
                                        st.session_state.past_df = df[~room_mask].copy()
                    
                    # Check for building name changes in existing rooms
                    for i, row in edited_df.iterrows():
                        if pd.isna(row['Building']) or pd.isna(row['Office']):
                            continue
                            
                        building = str(row['Building']).strip()
                        office = str(row['Office']).strip()
                        
                        # Find the original building for this office
                        original_rows = edit_df[edit_df['Office'] == office]
                        
                        for _, orig_row in original_rows.iterrows():
                            orig_building = orig_row['Building']
                            
                            # If building name changed, update all occupants
                            if orig_building != building:
                                # This is a building name change for an existing room
                                old_key = f"{orig_building}:{office}"
                                new_key = f"{building}:{office}"
                                
                                # Only process if the old key existed and new key is in our new set
                                if old_key in original_keys and new_key in new_keys and old_key != new_key:
                                    changes_made = True
                                    changes_summary.append(f"Renamed building for room {office}: {orig_building} → {building}")
                                    
                                    # Update all dataframes with the new building name
                                    dataframes = [
                                        st.session_state.current_df, 
                                        st.session_state.upcoming_df, 
                                        st.session_state.past_df
                                    ]
                                    
                                    for df in dataframes:
                                        if not df.empty:
                                            mask = (df['Building'] == orig_building) & (df['Office'] == office)
                                            if mask.any():
                                                df.loc[mask, 'Building'] = building
                    
                    # Save room capacities if changes were made
                    if changes_made:
                        save_room_capacities(st.session_state.room_capacities)
                        
                        # Show success message with summary of changes
                        st.success("Room information updated successfully!")
                        
                        if changes_summary:
                            with st.expander("View changes summary", expanded=True):
                                for change in changes_summary:
                                    st.write(f"- {change}")
                                
                                st.info("Remember to click 'Save Changes' in the sidebar to save all changes permanently.")
                        
                        # Refresh the page to show updated data
                        st.rerun()
                    else:
                        st.info("No changes detected in room information.")
                        
                except Exception as e:
                    st.error(f"Error updating room information: {e}")
                    st.exception(e)
        else:
            st.info("No room data available for editing")
            
            # Allow adding initial rooms if none exist
            with st.form("add_initial_room"):
                st.write("No rooms found. Add your first room:")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    new_building = st.text_input("Building Name", placeholder="e.g., Cockcroft")
                
                with col2:
                    new_office = st.text_input("Room Number", placeholder="e.g., 3.17")
                
                with col3:
                    new_capacity = st.number_input("Maximum Capacity", min_value=1, max_value=20, value=2)
                
                submitted = st.form_submit_button("Add First Room")
                
                if submitted:
                    if new_building and new_office:
                        # Add to the capacity dictionary
                        room_key = f"{new_building}:{new_office}"
                        st.session_state.room_capacities[room_key] = new_capacity
                        
                        # Add a placeholder in the current_df
                        placeholder = {
                            'Name': 'PLACEHOLDER',
                            'Status': 'Current',
                            'Email address': '',
                            'Position': '',
                            'Office': new_office,
                            'Building': new_building
                        }
                        temp_df = pd.DataFrame([placeholder])
                        st.session_state.current_df = pd.concat([st.session_state.current_df, temp_df], ignore_index=True)
                        
                        # Save and reload
                        save_room_capacities(st.session_state.room_capacities)
                        st.success(f"Added new room {new_office} in {new_building} with capacity {new_capacity}")
                        st.rerun()
                    else:
                        st.error("Building and Room Number are required")
    
    # Tab 3: Room Status
    with tab3:
        st.subheader("Room Status")
        
        # Create a view of all rooms and their current status with capacity information
        room_occupancy = get_room_occupancy_data()
        
        if not room_occupancy.empty:
            # Allow filtering
            building_filter = st.selectbox("Filter by Building", ['All'] + room_occupancy['Building'].unique().tolist(), key='status_building_filter')
            
            if building_filter != 'All':
                filtered_rooms = room_occupancy[room_occupancy['Building'] == building_filter]
            else:
                filtered_rooms = room_occupancy
            
            # Create status classes for rooms
            filtered_rooms['Status_Class'] = filtered_rooms.apply(
                lambda row: 'Storage' if row['IsStorage'] else
                           'Vacant' if row['Occupants'] == 0 else
                           'Low Occupancy' if row['Percentage'] <= 25 else
                           'Medium Occupancy' if row['Percentage'] <= 50 else
                           'High Occupancy' if row['Percentage'] <= 75 else
                           'Very High Occupancy' if row['Percentage'] < 100 else
                           'Full',
                axis=1
            )
            
            # Create status labels with occupancy info
            filtered_rooms['Status_Label'] = filtered_rooms.apply(
                lambda row: f"Storage" if row['IsStorage'] else
                           f"Vacant (0/{row['Max_Capacity']})" if row['Occupants'] == 0 else
                           f"Occupied ({row['Occupants']}/{row['Max_Capacity']} - {row['Percentage']:.1f}%)",
                axis=1
            )
            
            # Create display dataframe
            status_df = filtered_rooms[['Building', 'Floor', 'Office', 'Status_Class', 'Status_Label', 
                                         'Occupants', 'Max_Capacity', 'Remaining']]
            
            # Create a color mapping function for status
            def color_status(val):
                """
                Color cells based on status text.
                """
                color_map = {
                    'Storage': 'background-color: #e2e3e5',  # Gray for storage
                    'Vacant': 'background-color: #d4edda',  # Green for vacant
                    'Low Occupancy': 'background-color: #e6f7e1',  # Light green
                    'Medium Occupancy': 'background-color: #fff3cd',  # Yellow
                    'High Occupancy': 'background-color: #ffe5d9',  # Orange
                    'Very High Occupancy': 'background-color: #ffcccc',  # Pink
                    'Full': 'background-color: #f8d7da'  # Red
                }
                return color_map.get(val, '')

            # Use applymap to style the Status_Class column
            styled_status = status_df.style.applymap(
                lambda x: color_status(x) if isinstance(x, str) else '',
                subset=['Status_Class']
            )

            st.dataframe(styled_status, use_container_width=True)
            
            # Show status distribution
            status_counts = filtered_rooms['Status_Class'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            
            # Create a pie chart of room status
            fig = px.pie(
                status_counts, 
                values='Count', 
                names='Status',
                title="Room Status Distribution",
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            
            fig.update_traces(textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
            
            # Room availability by building
            st.subheader("Room Availability by Building")
            
            # Group by building and calculate availability metrics
            building_availability = filtered_rooms.groupby('Building').agg({
                'Office': 'count',
                'Occupants': 'sum',
                'Max_Capacity': 'sum',
                'Remaining': 'sum'
            }).reset_index()
            
            building_availability.rename(columns={'Office': 'Total Rooms'}, inplace=True)
            building_availability['Occupancy Rate'] = (building_availability['Occupants'] / 
                                                      building_availability['Max_Capacity'] * 100).round(1)
            
            st.dataframe(building_availability, use_container_width=True)
            
            # Stacked bar chart of occupancy by building
            fig = go.Figure()
            
            fig.add_trace(go.Bar(
                x=building_availability['Building'],
                y=building_availability['Occupants'],
                name='Occupied Spaces',
                marker_color='#4CAF50'
            ))
            
            fig.add_trace(go.Bar(
                x=building_availability['Building'],
                y=building_availability['Remaining'],
                name='Available Spaces',
                marker_color='#FFC107'
            ))
            
            fig.update_layout(
                barmode='stack',
                title='Capacity Utilization by Building',
                xaxis={'title': 'Building'},
                yaxis={'title': 'Number of Spaces'},
                legend={'title': 'Status'}
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No room status data available")
    
    # Tab 4: Room Assignment
    with tab4:
        st.subheader("Room Assignment Interface")
        
        # Get data for room occupancy
        room_occupancy = get_room_occupancy_data()
        
        if not room_occupancy.empty:
            # Create a two-column layout
            col1, col2 = st.columns([1, 3])
            
            with col1:
                st.write("**Select Person**")
                
                # Option to show current or upcoming people
                person_category = st.radio("Show", ["Current Occupants", "Upcoming Occupants"])
                
                if person_category == "Current Occupants":
                    people_df = st.session_state.current_df
                else:
                    people_df = st.session_state.upcoming_df
                
                # Filter to show people with and without room assignments
                if not people_df.empty:
                    # Consider someone unassigned if Office or Building is missing or empty
                    unassigned_mask = (
                        people_df['Office'].isna() | 
                        people_df['Building'].isna() |
                        (people_df['Office'] == '') | 
                        (people_df['Building'] == '')
                    )
                    
                    # Also exclude STORAGE records
                    not_storage_mask = ~people_df['Name'].str.contains('STORAGE', case=False, na=False)
                    
                    # Filter for assignment-eligible people
                    eligible_people = people_df[not_storage_mask].copy()

                    # Now add assignment status
                    eligible_people.loc[unassigned_mask & not_storage_mask, 'Assignment'] = "Unassigned"
                    eligible_people.loc[~unassigned_mask & not_storage_mask, 'Assignment'] = "Assigned"

                    # Create a selectbox for people, grouping by assignment status
                    assignment_filter = st.radio("Filter by", ["All", "Unassigned", "Assigned"])
                    
                    if assignment_filter != "All":
                        people_to_show = eligible_people[eligible_people['Assignment'] == assignment_filter]
                    else:
                        people_to_show = eligible_people
                    
                    if not people_to_show.empty:
                        # Format names to show assignment status
                        people_options = []
                        for _, person in people_to_show.iterrows():
                            name = person['Name']
                            assignment = person['Assignment']
                            display_name = f"{name} [{assignment}]"
                            people_options.append((display_name, name))
                        
                        # Create selectbox with formatted names
                        selected_display = st.selectbox(
                            "Select Person", 
                            [option[0] for option in people_options]
                        )
                        
                        # Extract actual name
                        selected_index = [option[0] for option in people_options].index(selected_display)
                        selected_person = people_options[selected_index][1]
                        
                        # Show current assignment if any
                        person_data = people_to_show[people_to_show['Name'] == selected_person].iloc[0]
                        if person_data['Assignment'] == "Assigned":
                            st.info(f"Currently assigned to: {person_data['Building']} - Room {person_data['Office']}")
                        else:
                            st.warning("Currently unassigned")
                        
                        # Show person details
                        details = {
                            "Name": person_data['Name'],
                            "Position": person_data.get('Position', 'Not specified'),
                            "Email": person_data.get('Email address', 'Not specified')
                        }
                        
                        for label, value in details.items():
                            if pd.notna(value) and value:
                                st.write(f"**{label}:** {value}")
                    else:
                        st.info(f"No {assignment_filter.lower()} {person_category.lower()} found")
                        selected_person = None
                else:
                    st.info(f"No {person_category.lower()} data available")
                    selected_person = None
            
            with col2:
                st.write("**Available Rooms**")
                
                # Filter controls for rooms
                col_a, col_b, col_c = st.columns(3)
                with col_a:
                    building_filter = st.selectbox("Building", ['All'] + room_occupancy['Building'].unique().tolist())
                with col_b:
                    floor_filter = st.selectbox("Floor", ['All'] + sorted(room_occupancy['Floor'].unique()))
                with col_c:
                    capacity_filter = st.multiselect(
                        "Show Rooms", 
                        ["Vacant", "Has Space", "Full", "Storage"],
                        default=["Vacant", "Has Space"]
                    )
                
                # Apply filters
                filtered_rooms = room_occupancy.copy()
                if building_filter != 'All':
                    filtered_rooms = filtered_rooms[filtered_rooms['Building'] == building_filter]
                
                if floor_filter != 'All':
                    filtered_rooms = filtered_rooms[filtered_rooms['Floor'] == floor_filter]
                
                # Apply capacity filters
                capacity_conditions = []
                if "Vacant" in capacity_filter:
                    capacity_conditions.append(filtered_rooms['Occupants'] == 0)
                if "Has Space" in capacity_filter:
                    capacity_conditions.append((filtered_rooms['Occupants'] > 0) & (filtered_rooms['Remaining'] > 0))
                if "Full" in capacity_filter:
                    capacity_conditions.append(filtered_rooms['Remaining'] <= 0)
                if "Storage" in capacity_filter:
                    capacity_conditions.append(filtered_rooms['IsStorage'])
                
                if capacity_conditions:
                    combined_condition = capacity_conditions[0]
                    for condition in capacity_conditions[1:]:
                        combined_condition = combined_condition | condition
                    filtered_rooms = filtered_rooms[combined_condition]
                
                # Order by remaining capacity (most available first)
                filtered_rooms = filtered_rooms.sort_values(['Building', 'Floor', 'Remaining'], ascending=[True, True, False])
                
                if not filtered_rooms.empty:
                    st.write(f"Showing {len(filtered_rooms)} rooms matching your filters")
                    
                    # Create a grid layout of room cards
                    rooms_per_row = 3
                    num_rooms = len(filtered_rooms)
                    num_rows = (num_rooms + rooms_per_row - 1) // rooms_per_row
                    
                    for row in range(num_rows):
                        cols = st.columns(rooms_per_row)
                        for col in range(rooms_per_row):
                            idx = row * rooms_per_row + col
                            if idx < num_rooms:
                                room = filtered_rooms.iloc[idx]
                                
                                # Get room details
                                building = room['Building']
                                office = room['Office']
                                occupants = room['Occupants']
                                max_capacity = room['Max_Capacity']
                                remaining = room['Remaining']
                                percentage = room['Percentage']
                                is_storage = room['IsStorage']
                                
                                # Get occupant information
                                room_occupants = []
                                if occupants > 0:
                                    occupant_data = st.session_state.current_df[
                                        (st.session_state.current_df['Building'] == building) & 
                                        (st.session_state.current_df['Office'] == office)
                                    ]
                                    room_occupants = occupant_data['Name'].tolist()
                                
                                # Determine room status
                                if is_storage:
                                    status_class = "room-storage"
                                    status_text = "Storage Room"
                                elif occupants == 0:
                                    status_class = "room-vacant"
                                    status_text = f"Vacant ({remaining} available)"
                                elif remaining > 0:
                                    if percentage <= 25:
                                        status_class = "room-low"
                                    elif percentage <= 50:
                                        status_class = "room-medium"
                                    elif percentage <= 75:
                                        status_class = "room-high"
                                    else:
                                        status_class = "room-high"
                                    status_text = f"Has Space ({remaining} available)"
                                else:
                                    status_class = "room-full"
                                    status_text = f"Full ({occupants}/{max_capacity})"
                                
                                with cols[col]:
                                    # Create room card
                                    st.markdown(
                                        f"""
                                        <div class="{status_class}">
                                            <h4 style="margin:0;">{building} - {office}</h4>
                                            <p style="margin:0;">{status_text}</p>
                                            <p style="margin:0;">Occupancy: {occupants}/{max_capacity} ({percentage:.1f}%)</p>
                                        </div>
                                        """, 
                                        unsafe_allow_html=True
                                    )
                                    
                                    # Show current occupants
                                    if room_occupants:
                                        st.markdown("**Current occupants:**")
                                        for occupant in room_occupants:
                                            if "STORAGE" not in occupant:
                                                st.markdown(f"- {occupant}")
                                    
                                    # Show assign button
                                    if selected_person and not is_storage:
                                        # Only allow assignment if room has space
                                        if remaining > 0:
                                            if st.button(f"Assign to Room {office}", key=f"assign_{building}_{office}"):
                                                # Find the person in the appropriate dataframe
                                                if person_category == "Current Occupants":
                                                    target_df = st.session_state.current_df
                                                else:
                                                    target_df = st.session_state.upcoming_df
                                                
                                                # Find the person's row
                                                person_idx = target_df[target_df['Name'] == selected_person].index
                                                
                                                if len(person_idx) > 0:
                                                    # Update the person's room assignment
                                                    target_df.loc[person_idx[0], 'Building'] = building
                                                    target_df.loc[person_idx[0], 'Office'] = office
                                                    
                                                    st.success(f"✅ Assigned {selected_person} to {building} - Room {office}. Remember to save changes!")
                                                    
                                                    # Force a rerun to update the UI
                                                    st.rerun()
                                        else:
                                            st.error("Room is full")
                else:
                    st.info("No rooms match the selected filters")
        else:
            st.info("No room data available for assignment")

# Reports Page
elif page == "Reports":
    st.title("Office Allocation Reports")
    
    # Create tabs for different report types
    report_tabs = st.tabs([
        "Occupancy Summary", 
        "Building Reports", 
        "Room Utilization", 
        "Occupant Reports",
        "Export Data"
    ])
    
    # Tab 1: Occupancy Summary
    with report_tabs[0]:
        st.subheader("Occupancy Summary Report")
        
        # Display date range for the report
        col1, col2 = st.columns(2)
        with col1:
            report_date = st.date_input("Report Date", datetime.now())
        with col2:
            st.metric("Data Last Updated", 
                     st.session_state.last_save if st.session_state.last_save else "Not saved yet")
        
        # Get summary metrics
        summary_metrics = {
            "Total Buildings": len(get_unique_buildings()),
            "Total Rooms": len(get_unique_offices()),
            "Current Occupants": len(st.session_state.current_df),
            "Upcoming Occupants": len(st.session_state.upcoming_df),
            "Past Occupants": len(st.session_state.past_df)
        }
        
        # Add occupancy metrics if we have room data
        room_occupancy = get_room_occupancy_data()
        if not room_occupancy.empty:
            summary_metrics.update({
                "Total Capacity": room_occupancy['Max_Capacity'].sum(),
                "Currently Occupied": room_occupancy['Occupants'].sum(),
                "Available Spaces": room_occupancy['Remaining'].sum(),
                "Occupancy Rate": f"{(room_occupancy['Occupants'].sum() / room_occupancy['Max_Capacity'].sum() * 100):.1f}%" if room_occupancy['Max_Capacity'].sum() > 0 else "0%"
            })
        
        # Display summary metrics in a nice format
        st.markdown("### Key Metrics")
        
        # Use columns to display metrics in rows of 3
        metrics = list(summary_metrics.items())
        for i in range(0, len(metrics), 3):
            cols = st.columns(3)
            for j in range(3):
                if i + j < len(metrics):
                    key, value = metrics[i + j]
                    cols[j].metric(key, value)
        
        # Display occupancy trends if we have room data
        if not room_occupancy.empty:
            st.markdown("### Occupancy by Building")
            
            # Group by building
            building_occupancy = room_occupancy.groupby('Building').agg({
                'Office': 'count',
                'Occupants': 'sum',
                'Max_Capacity': 'sum',
                'Remaining': 'sum'
            }).reset_index()
            
            building_occupancy.rename(columns={'Office': 'Room Count'}, inplace=True)
            building_occupancy['Occupancy Rate'] = (building_occupancy['Occupants'] / 
                                                 building_occupancy['Max_Capacity'] * 100).round(1)
            
            # Create stacked bar chart
            fig = go.Figure()
            
            fig.add_trace(go.Bar(
                x=building_occupancy['Building'],
                y=building_occupancy['Occupants'],
                name='Occupied',
                marker_color='#4CAF50',
                text=building_occupancy['Occupants'],
                textposition='auto'
            ))
            
            fig.add_trace(go.Bar(
                x=building_occupancy['Building'],
                y=building_occupancy['Remaining'],
                name='Available',
                marker_color='#FFC107',
                text=building_occupancy['Remaining'],
                textposition='auto'
            ))
            
            fig.update_layout(
                barmode='stack',
                title='Building Occupancy',
                xaxis_title='Building',
                yaxis_title='Number of Places',
                legend_title='Status'
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Comparison with upcoming occupants
            st.markdown("### Current vs. Upcoming Occupancy")
            
            # Count upcoming occupants by building
            if not st.session_state.upcoming_df.empty and 'Building' in st.session_state.upcoming_df.columns:
                upcoming_by_building = st.session_state.upcoming_df['Building'].value_counts().reset_index()
                upcoming_by_building.columns = ['Building', 'Upcoming']
                
                # Merge with current occupancy
                merged_occupancy = building_occupancy.merge(
                    upcoming_by_building, 
                    on='Building', 
                    how='left'
                ).fillna(0)
                
                # Create side-by-side bar chart
                fig = go.Figure()
                
                fig.add_trace(go.Bar(
                    x=merged_occupancy['Building'],
                    y=merged_occupancy['Occupants'],
                    name='Current',
                    marker_color='#4CAF50',
                    text=merged_occupancy['Occupants'].astype(int),
                    textposition='auto'
                ))
                
                fig.add_trace(go.Bar(
                    x=merged_occupancy['Building'],
                    y=merged_occupancy['Upcoming'],
                    name='Upcoming',
                    marker_color='#2196F3',
                    text=merged_occupancy['Upcoming'].astype(int),
                    textposition='auto'
                ))
                
                fig.update_layout(
                    barmode='group',
                    title='Current vs. Upcoming Occupants by Building',
                    xaxis_title='Building',
                    yaxis_title='Number of Occupants',
                    legend_title='Status'
                )
                
                st.plotly_chart(fig, use_container_width=True)
    
    # Tab 2: Building Reports
    with report_tabs[1]:
        st.subheader("Building Reports")
        
        # Select building to report on
        selected_building = st.selectbox(
            "Select Building", 
            get_unique_buildings(),
            key="report_building_select"
        )
        
        if selected_building:
            # Filter room data for this building
            building_rooms = room_occupancy[room_occupancy['Building'] == selected_building].copy()
            
            if not building_rooms.empty:
                # Building summary
                st.markdown(f"### {selected_building} Summary")
                
                # Calculate building metrics
                total_rooms = len(building_rooms)
                total_capacity = building_rooms['Max_Capacity'].sum()
                total_occupants = building_rooms['Occupants'].sum()
                available_spaces = building_rooms['Remaining'].sum()
                occupancy_rate = (total_occupants / total_capacity * 100) if total_capacity > 0 else 0
                
                # Display metrics in a row
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total Rooms", total_rooms)
                col2.metric("Total Capacity", total_capacity)
                col3.metric("Current Occupants", total_occupants)
                col4.metric("Occupancy Rate", f"{occupancy_rate:.1f}%")
                
                # Show floors in this building
                st.markdown("### Floors Overview")
                
                # Group by floor
                floor_data = building_rooms.groupby('Floor').agg({
                    'Office': 'count',
                    'Occupants': 'sum',
                    'Max_Capacity': 'sum',
                    'Remaining': 'sum'
                }).reset_index()
                
                floor_data.rename(columns={'Office': 'Room Count'}, inplace=True)
                floor_data['Occupancy Rate'] = (floor_data['Occupants'] / floor_data['Max_Capacity'] * 100).round(1)
                
                # Sort by floor number
                floor_data['Floor_Num'] = floor_data['Floor'].apply(
                    lambda x: float(x) if x.replace('.', '', 1).isdigit() else float('inf')
                )
                floor_data = floor_data.sort_values('Floor_Num').drop('Floor_Num', axis=1)
                
                # Display as a table
                st.dataframe(floor_data, use_container_width=True)
                
                # Create a visualization of rooms by floor
                st.markdown("### Room Occupancy by Floor")
                
                # Create a stacked bar chart of room occupancy by floor
                fig = go.Figure()
                
                fig.add_trace(go.Bar(
                    x=floor_data['Floor'],
                    y=floor_data['Occupants'],
                    name='Occupied',
                    marker_color='#4CAF50',
                    text=floor_data['Occupants'].astype(int),
                    textposition='auto'
                ))
                
                fig.add_trace(go.Bar(
                    x=floor_data['Floor'],
                    y=floor_data['Remaining'],
                    name='Available',
                    marker_color='#FFC107',
                    text=floor_data['Remaining'].astype(int),
                    textposition='auto'
                ))
                
                fig.update_layout(
                    barmode='stack',
                    title=f'Room Occupancy by Floor in {selected_building}',
                    xaxis_title='Floor',
                    yaxis_title='Number of Places',
                    legend_title='Status'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # List all occupants in this building
                st.markdown("### Current Occupants")
                
                # Filter current occupants for this building
                building_occupants = st.session_state.current_df[
                    st.session_state.current_df['Building'] == selected_building
                ].copy()
                
                if not building_occupants.empty:
                    # Sort by Office (Room) for better organization
                    building_occupants = building_occupants.sort_values(['Office', 'Name'])
                    
                    # Display table of occupants
                    st.dataframe(
                        building_occupants[['Name', 'Position', 'Office', 'Email address']],
                        use_container_width=True
                    )
                    
                    # Count occupants by position
                    if 'Position' in building_occupants.columns:
                        position_counts = building_occupants['Position'].value_counts()
                        
                        # Create pie chart of positions
                        if len(position_counts) > 0:
                            st.markdown("### Occupants by Position")
                            fig = px.pie(
                                values=position_counts.values,
                                names=position_counts.index,
                                title=f"Positions in {selected_building}"
                            )
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info(f"No current occupants in {selected_building}")
            else:
                st.warning(f"No room data available for {selected_building}")
    
    # Tab 3: Room Utilization
    with report_tabs[2]:
        st.subheader("Room Utilization Report")
        
        if not room_occupancy.empty:
            # Create room utilization categories
            room_occupancy['Utilization'] = room_occupancy.apply(
                lambda row: 'Storage' if row['IsStorage'] else
                           'Vacant' if row['Occupants'] == 0 else
                           'Low (1-25%)' if row['Percentage'] <= 25 else
                           'Medium (26-50%)' if row['Percentage'] <= 50 else
                           'High (51-75%)' if row['Percentage'] <= 75 else
                           'Very High (76-99%)' if row['Percentage'] < 100 else
                           'Full (100%)',
                axis=1
            )
            
            # Count rooms by utilization category
            util_counts = room_occupancy['Utilization'].value_counts().reset_index()
            util_counts.columns = ['Utilization', 'Count']
            
            # Order categories
            category_order = ['Vacant', 'Low (1-25%)', 'Medium (26-50%)', 'High (51-75%)', 
                             'Very High (76-99%)', 'Full (100%)', 'Storage']
            util_counts['Utilization'] = pd.Categorical(
                util_counts['Utilization'], 
                categories=category_order, 
                ordered=True
            )
            util_counts = util_counts.sort_values('Utilization')
            
            # Create color map for utilization categories
            color_map = {
                'Vacant': '#d4edda',
                'Low (1-25%)': '#e6f7e1',
                'Medium (26-50%)': '#fff3cd',
                'High (51-75%)': '#ffe5d9',
                'Very High (76-99%)': '#ffcccc',
                'Full (100%)': '#f8d7da',
                'Storage': '#e2e3e5'
            }
            
            # Create bar chart of room utilization
            fig = px.bar(
                util_counts,
                x='Utilization',
                y='Count',
                color='Utilization',
                color_discrete_map=color_map,
                text='Count',
                title='Room Utilization Distribution'
            )
            
            fig.update_layout(
                xaxis_title='Utilization Category',
                yaxis_title='Number of Rooms',
                showlegend=False
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # List rooms that need attention (high utilization or vacant)
            st.markdown("### Rooms Requiring Attention")
            
            # Create tabs for different categories
            attention_tabs = st.tabs([
                "Full Rooms", 
                "Very High Utilization", 
                "Vacant Rooms"
            ])
            
            # Tab 1: Full Rooms
            with attention_tabs[0]:
                full_rooms = room_occupancy[room_occupancy['Utilization'] == 'Full (100%)'].copy()
                
                if not full_rooms.empty:
                    st.warning(f"There are {len(full_rooms)} rooms at full capacity")
                    
                    # Display the full rooms
                    st.dataframe(
                        full_rooms[['Building', 'Floor', 'Office', 'Occupants', 'Max_Capacity']],
                        use_container_width=True
                    )
                    
                    # For each full room, show occupants
                    st.markdown("### Occupants of Full Rooms")
                    
                    for _, room in full_rooms.iterrows():
                        building = room['Building']
                        office = room['Office']
                        
                        # Get occupants for this room
                        room_occupants = st.session_state.current_df[
                            (st.session_state.current_df['Building'] == building) &
                            (st.session_state.current_df['Office'] == office)
                        ]
                        
                        if not room_occupants.empty:
                            st.markdown(f"**{building} - Room {office}**")
                            
                            for _, occupant in room_occupants.iterrows():
                                st.markdown(f"- {occupant['Name']} ({occupant.get('Position', 'No position')})")
                else:
                    st.success("No rooms are currently at full capacity")
            
            # Tab 2: Very High Utilization
            with attention_tabs[1]:
                high_util_rooms = room_occupancy[room_occupancy['Utilization'] == 'Very High (76-99%)'].copy()
                
                if not high_util_rooms.empty:
                    st.warning(f"There are {len(high_util_rooms)} rooms with very high utilization (76-99%)")
                    
                    # Display the high utilization rooms
                    high_util_rooms['Remaining_Places'] = high_util_rooms['Max_Capacity'] - high_util_rooms['Occupants']
                    
                    st.dataframe(
                        high_util_rooms[['Building', 'Floor', 'Office', 'Occupants', 
                                       'Max_Capacity', 'Remaining_Places', 'Percentage']],
                        use_container_width=True
                    )
                else:
                    st.success("No rooms currently have very high utilization")
            
            # Tab 3: Vacant Rooms
            with attention_tabs[2]:
                vacant_rooms = room_occupancy[room_occupancy['Utilization'] == 'Vacant'].copy()
                
                if not vacant_rooms.empty:
                    st.info(f"There are {len(vacant_rooms)} vacant rooms that could be utilized")
                    
                    # Display the vacant rooms
                    vacant_rooms = vacant_rooms.sort_values(['Building', 'Floor', 'Office'])
                    
                    st.dataframe(
                        vacant_rooms[['Building', 'Floor', 'Office', 'Max_Capacity']],
                        use_container_width=True
                    )
                    
                    # Show distribution of vacant rooms by building
                    vacant_by_building = vacant_rooms.groupby('Building').size().reset_index(name='Vacant Rooms')
                    
                    fig = px.bar(
                        vacant_by_building,
                        x='Building',
                        y='Vacant Rooms',
                        title='Vacant Rooms by Building',
                        text='Vacant Rooms'
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("All rooms are currently occupied")
            
            # Add a summary table of all rooms
            st.markdown("### Complete Room Utilization Table")
            
            # Create a formatted table with all rooms
            room_table = room_occupancy.sort_values(['Building', 'Floor', 'Office'])
            
            # Display the table
            st.dataframe(
                room_table[['Building', 'Floor', 'Office', 'Occupants', 
                         'Max_Capacity', 'Remaining', 'Percentage', 'Utilization']],
                use_container_width=True
            )
        else:
            st.info("No room occupancy data available")
    
    # Tab 4: Occupant Reports
    with report_tabs[3]:
        st.subheader("Occupant Reports")
        
        # Create tabs for occupant reports
        occupant_tabs = st.tabs([
            "Current Occupants", 
            "Upcoming Occupants", 
            "Past Occupants", 
            "Position Analysis"
        ])
        
        # Tab 1: Current Occupants
        with occupant_tabs[0]:
            if not st.session_state.current_df.empty:
                st.markdown(f"### Current Occupants ({len(st.session_state.current_df)})")
                
                # Allow filtering
                filter_options = st.multiselect(
                    "Filter by Building", 
                    get_unique_buildings(),
                    key="curr_occupant_filter"
                )
                
                # Filter data if needed
                if filter_options:
                    filtered_occupants = st.session_state.current_df[
                        st.session_state.current_df['Building'].isin(filter_options)
                    ]
                else:
                    filtered_occupants = st.session_state.current_df
                
                # Sort by building and room for better organization
                filtered_occupants = filtered_occupants.sort_values(['Building', 'Office', 'Name'])
                
                # Display the table
                st.dataframe(
                    filtered_occupants[['Name', 'Position', 'Building', 'Office', 'Email address']],
                    use_container_width=True
                )
                
                # Create a summary chart
                occupants_by_building = filtered_occupants['Building'].value_counts().reset_index()
                occupants_by_building.columns = ['Building', 'Occupants']
                
                fig = px.bar(
                    occupants_by_building,
                    x='Building',
                    y='Occupants',
                    title='Current Occupants by Building',
                    text='Occupants'
                )
                
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No current occupants data available")
        
        # Tab 2: Upcoming Occupants
        with occupant_tabs[1]:
            if not st.session_state.upcoming_df.empty:
                st.markdown(f"### Upcoming Occupants ({len(st.session_state.upcoming_df)})")
                
                # Display the table
                st.dataframe(
                    st.session_state.upcoming_df[['Name', 'Position', 'Building', 'Office', 'Email address']],
                    use_container_width=True
                )
                
                # Create a summary chart if we have building information
                if 'Building' in st.session_state.upcoming_df.columns:
                    upcoming_by_building = st.session_state.upcoming_df['Building'].value_counts().reset_index()
                    upcoming_by_building.columns = ['Building', 'Upcoming']
                    fig = px.bar(
                        upcoming_by_building,
                        x='Building',
                        y='Upcoming',
                        title='Upcoming Occupants by Building',
                        text='Upcoming'
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No upcoming occupants data available")
        
        # Tab 3: Past Occupants
        with occupant_tabs[2]:
            if not st.session_state.past_df.empty:
                st.markdown(f"### Past Occupants ({len(st.session_state.past_df)})")
                
                # Add search functionality
                search_term = st.text_input("Search by name, position, or email", key="past_report_search")
                
                filtered_past = st.session_state.past_df
                
                if search_term:
                    filtered_past = filtered_past[
                        filtered_past['Name'].str.contains(search_term, case=False, na=False) |
                        filtered_past['Position'].str.contains(search_term, case=False, na=False) |
                        filtered_past['Email address'].str.contains(search_term, case=False, na=False)
                    ]
                
                # Display the table
                st.dataframe(
                    filtered_past[['Name', 'Position', 'Building', 'Office', 'Email address']],
                    use_container_width=True
                )
                
                # Show timeline of departures if we have date information
                if 'End Date' in filtered_past.columns:
                    st.markdown("### Departures Timeline")
                    
                    # Count departures by month
                    filtered_past['End_Month'] = pd.to_datetime(filtered_past['End Date']).dt.strftime('%Y-%m')
                    departures_by_month = filtered_past['End_Month'].value_counts().sort_index().reset_index()
                    departures_by_month.columns = ['Month', 'Departures']
                    
                    fig = px.line(
                        departures_by_month,
                        x='Month',
                        y='Departures',
                        title='Past Occupants Departure Timeline',
                        markers=True
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No past occupants data available")
        
        # Tab 4: Position Analysis
        with occupant_tabs[3]:
            if not st.session_state.current_df.empty and 'Position' in st.session_state.current_df.columns:
                st.markdown("### Occupant Position Analysis")
                
                # Get position counts
                position_counts = st.session_state.current_df['Position'].fillna('Not Specified').value_counts()
                
                # If there are too many positions, group smaller ones
                if len(position_counts) > 8:
                    top_positions = position_counts.head(7)
                    other_count = position_counts.tail(len(position_counts) - 7).sum()
                    position_counts = pd.concat([top_positions, pd.Series([other_count], index=['Other'])])
                
                # Create dataframe for visualization
                position_df = position_counts.reset_index()
                position_df.columns = ['Position', 'Count']
                
                # Create pie chart
                fig = px.pie(
                    position_df,
                    values='Count',
                    names='Position',
                    title='Current Occupants by Position'
                )
                
                fig.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig, use_container_width=True)
                
                # Show positions by building
                st.markdown("### Positions by Building")
                
                # Create a crosstab of building vs position
                building_position = pd.crosstab(
                    st.session_state.current_df['Building'], 
                    st.session_state.current_df['Position'].fillna('Not Specified')
                ).reset_index()
                
                # Display the crosstab
                st.dataframe(building_position, use_container_width=True)
                
                # Create a stacked bar chart
                position_building_data = []
                
                for position in position_df['Position']:
                    if position in building_position.columns:
                        for building in building_position['Building']:
                            position_count = building_position.loc[
                                building_position['Building'] == building, 
                                position
                            ].values[0]
                            
                            position_building_data.append({
                                'Building': building,
                                'Position': position,
                                'Count': position_count
                            })
                
                if position_building_data:
                    position_building_df = pd.DataFrame(position_building_data)
                    
                    fig = px.bar(
                        position_building_df,
                        x='Building',
                        y='Count',
                        color='Position',
                        title='Positions by Building',
                        barmode='stack'
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No position data available for analysis")
    
    # Tab 5: Export Data
    with report_tabs[4]:
        st.subheader("Export Data")
        
        # Create tabs for different export options
        export_tabs = st.tabs([
            "CSV Export", 
            "Excel Reports", 
            "Building Summary"
        ])
        
        # Tab 1: CSV Export
        with export_tabs[0]:
            st.markdown("### Export Data as CSV")
            
            # Create export options
            export_options = st.multiselect(
                "Select data to export",
                ["Current Occupants", "Upcoming Occupants", "Past Occupants", "Room Utilization"],
                default=["Current Occupants"]
            )
            
            if st.button("Generate CSV Files"):
                # Create a zip file with CSVs
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                export_dir = f"data/export_{timestamp}"
                os.makedirs(export_dir, exist_ok=True)
                
                exported_files = []
                
                # Export selected dataframes
                if "Current Occupants" in export_options and not st.session_state.current_df.empty:
                    current_file = f"{export_dir}/current_occupants.csv"
                    st.session_state.current_df.to_csv(current_file, index=False)
                    exported_files.append(("Current Occupants", current_file))
                
                if "Upcoming Occupants" in export_options and not st.session_state.upcoming_df.empty:
                    upcoming_file = f"{export_dir}/upcoming_occupants.csv"
                    st.session_state.upcoming_df.to_csv(upcoming_file, index=False)
                    exported_files.append(("Upcoming Occupants", upcoming_file))
                
                if "Past Occupants" in export_options and not st.session_state.past_df.empty:
                    past_file = f"{export_dir}/past_occupants.csv"
                    st.session_state.past_df.to_csv(past_file, index=False)
                    exported_files.append(("Past Occupants", past_file))
                
                if "Room Utilization" in export_options:
                    room_occupancy = get_room_occupancy_data()
                    if not room_occupancy.empty:
                        rooms_file = f"{export_dir}/room_utilization.csv"
                        room_occupancy.to_csv(rooms_file, index=False)
                        exported_files.append(("Room Utilization", rooms_file))
                
                # Display download links
                if exported_files:
                    st.success(f"Generated {len(exported_files)} CSV files")
                    
                    for name, file_path in exported_files:
                        with open(file_path, "rb") as file:
                            st.download_button(
                                label=f"Download {name} CSV",
                                data=file,
                                file_name=os.path.basename(file_path),
                                mime="text/csv",
                                key=f"dl_{name}"
                            )
                else:
                    st.error("No files were exported. Please select data to export.")
        
        # Tab 2: Excel Reports
        with export_tabs[1]:
            st.markdown("### Generate Excel Report")
            
            # Create report options
            report_type = st.radio(
                "Report Type",
                ["Full Office Allocation Report", "Building-Specific Report", "Utilization Summary"]
            )
            
            if report_type == "Building-Specific Report":
                report_building = st.selectbox(
                    "Select Building",
                    get_unique_buildings(),
                    key="excel_report_building"
                )
            
            if st.button("Generate Excel Report"):
                # Create Excel file
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                if report_type == "Full Office Allocation Report":
                    report_file = f"data/Full_Office_Report_{timestamp}.xlsx"
                    
                    # Create a writer
                    with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
                        # Current occupants
                        if not st.session_state.current_df.empty:
                            st.session_state.current_df.to_excel(writer, sheet_name='Current Occupants', index=False)
                        
                        # Upcoming occupants
                        if not st.session_state.upcoming_df.empty:
                            st.session_state.upcoming_df.to_excel(writer, sheet_name='Upcoming Occupants', index=False)
                        
                        # Room data
                        room_occupancy = get_room_occupancy_data()
                        if not room_occupancy.empty:
                            room_occupancy.to_excel(writer, sheet_name='Room Utilization', index=False)
                        
                        # Summary sheet
                        summary_data = [
                            ["Office Allocation Report", "", ""],
                            ["Generated on", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""],
                            ["", "", ""],
                            ["Metric", "Value", ""],
                            ["Total Buildings", len(get_unique_buildings()), ""],
                            ["Total Rooms", len(get_unique_offices()), ""],
                            ["Current Occupants", len(st.session_state.current_df), ""],
                            ["Upcoming Occupants", len(st.session_state.upcoming_df), ""],
                            ["Past Occupants", len(st.session_state.past_df), ""]
                        ]
                        
                        # Add occupancy metrics if we have room data
                        if not room_occupancy.empty:
                            occupancy_rate = (room_occupancy['Occupants'].sum() / 
                                             room_occupancy['Max_Capacity'].sum() * 100) if room_occupancy['Max_Capacity'].sum() > 0 else 0
                            
                            summary_data.extend([
                                ["Total Capacity", room_occupancy['Max_Capacity'].sum(), ""],
                                ["Currently Occupied", room_occupancy['Occupants'].sum(), ""],
                                ["Available Spaces", room_occupancy['Remaining'].sum(), ""],
                                ["Occupancy Rate", f"{occupancy_rate:.1f}%", ""]
                            ])
                        
                        # Create summary sheet
                        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False, header=False)
                    
                    # Provide download link
                    with open(report_file, "rb") as file:
                        st.download_button(
                            label="Download Full Report",
                            data=file,
                            file_name=os.path.basename(report_file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                elif report_type == "Building-Specific Report" and report_building:
                    report_file = f"data/{report_building}_Report_{timestamp}.xlsx"
                    
                    # Create a writer
                    with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
                        # Filter data for this building
                        current_building = st.session_state.current_df[
                            st.session_state.current_df['Building'] == report_building
                        ] if not st.session_state.current_df.empty else pd.DataFrame()
                        
                        upcoming_building = st.session_state.upcoming_df[
                            st.session_state.upcoming_df['Building'] == report_building
                        ] if not st.session_state.upcoming_df.empty else pd.DataFrame()
                        
                        room_occupancy = get_room_occupancy_data()
                        building_rooms = room_occupancy[
                            room_occupancy['Building'] == report_building
                        ] if not room_occupancy.empty else pd.DataFrame()
                        
                        # Write sheets
                        if not current_building.empty:
                            current_building.to_excel(writer, sheet_name='Current Occupants', index=False)
                        
                        if not upcoming_building.empty:
                            upcoming_building.to_excel(writer, sheet_name='Upcoming Occupants', index=False)
                        
                        if not building_rooms.empty:
                            building_rooms.to_excel(writer, sheet_name='Rooms', index=False)
                        
                        # Summary sheet
                        total_rooms = len(building_rooms) if not building_rooms.empty else 0
                        total_capacity = building_rooms['Max_Capacity'].sum() if not building_rooms.empty else 0
                        total_occupants = building_rooms['Occupants'].sum() if not building_rooms.empty else 0
                        available_spaces = building_rooms['Remaining'].sum() if not building_rooms.empty else 0
                        occupancy_rate = (total_occupants / total_capacity * 100) if total_capacity > 0 else 0
                        
                        summary_data = [
                            [f"{report_building} Building Report", "", ""],
                            ["Generated on", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""],
                            ["", "", ""],
                            ["Metric", "Value", ""],
                            ["Total Rooms", total_rooms, ""],
                            ["Total Capacity", total_capacity, ""],
                            ["Current Occupants", total_occupants, ""],
                            ["Available Spaces", available_spaces, ""],
                            ["Occupancy Rate", f"{occupancy_rate:.1f}%", ""],
                            ["Upcoming Occupants", len(upcoming_building), ""]
                        ]
                        
                        # Create summary sheet
                        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False, header=False)
                    
                    # Provide download link
                    with open(report_file, "rb") as file:
                        st.download_button(
                            label=f"Download {report_building} Report",
                            data=file,
                            file_name=os.path.basename(report_file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                elif report_type == "Utilization Summary":
                    report_file = f"data/Utilization_Summary_{timestamp}.xlsx"
                    
                    # Create room occupancy data with categories
                    room_occupancy = get_room_occupancy_data()
                    
                    if not room_occupancy.empty:
                        # Add utilization category
                        room_occupancy['Utilization'] = room_occupancy.apply(
                            lambda row: 'Storage' if row['IsStorage'] else
                                      'Vacant' if row['Occupants'] == 0 else
                                      'Low (1-25%)' if row['Percentage'] <= 25 else
                                      'Medium (26-50%)' if row['Percentage'] <= 50 else
                                      'High (51-75%)' if row['Percentage'] <= 75 else
                                      'Very High (76-99%)' if row['Percentage'] < 100 else
                                      'Full (100%)',
                            axis=1
                        )
                        
                        # Create a writer
                        with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
                            # All rooms with utilization
                            room_occupancy.to_excel(writer, sheet_name='All Rooms', index=False)
                            
                            # Full rooms
                            full_rooms = room_occupancy[room_occupancy['Utilization'] == 'Full (100%)']
                            if not full_rooms.empty:
                                full_rooms.to_excel(writer, sheet_name='Full Rooms', index=False)
                            
                            # Vacant rooms
                            vacant_rooms = room_occupancy[room_occupancy['Utilization'] == 'Vacant']
                            if not vacant_rooms.empty:
                                vacant_rooms.to_excel(writer, sheet_name='Vacant Rooms', index=False)
                            
                            # High utilization rooms
                            high_util = room_occupancy[room_occupancy['Utilization'] == 'Very High (76-99%)']
                            if not high_util.empty:
                                high_util.to_excel(writer, sheet_name='High Utilization', index=False)
                            
                            # Summary by building
                            building_summary = room_occupancy.groupby('Building').agg({
                                'Office': 'count',
                                'Occupants': 'sum',
                                'Max_Capacity': 'sum',
                                'Remaining': 'sum'
                            }).reset_index()
                            
                            building_summary.rename(columns={'Office': 'Room Count'}, inplace=True)
                            building_summary['Occupancy Rate'] = (building_summary['Occupants'] / 
                                                              building_summary['Max_Capacity'] * 100).round(1)
                            
                            building_summary.to_excel(writer, sheet_name='Building Summary', index=False)
                        
                        # Provide download link
                        with open(report_file, "rb") as file:
                            st.download_button(
                                label="Download Utilization Report",
                                data=file,
                                file_name=os.path.basename(report_file),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("No room data available to generate utilization report")
        
        # Tab 3: Building Summary
        with export_tabs[2]:
            st.markdown("### Building Summary Report")
            
            # Generate a summary of all buildings
            room_occupancy = get_room_occupancy_data()
            
            if not room_occupancy.empty:
                # Create building summary
                building_summary = room_occupancy.groupby('Building').agg({
                    'Office': 'count',
                    'Occupants': 'sum',
                    'Max_Capacity': 'sum',
                    'Remaining': 'sum'
                }).reset_index()
                
                building_summary.rename(columns={'Office': 'Room Count'}, inplace=True)
                building_summary['Occupancy Rate'] = (building_summary['Occupants'] / 
                                                  building_summary['Max_Capacity'] * 100).round(1)
                
                # Display summary table
                st.dataframe(building_summary, use_container_width=True)
                
                # Create a visualization
                fig = make_subplots(specs=[[{"secondary_y": True}]])
                
                # Add bars for room count
                fig.add_trace(
                    go.Bar(
                        x=building_summary['Building'],
                        y=building_summary['Room Count'],
                        name='Room Count',
                        marker_color='#4CAF50'
                    ),
                    secondary_y=False
                )
                
                # Add line for occupancy rate
                fig.add_trace(
                    go.Scatter(
                        x=building_summary['Building'],
                        y=building_summary['Occupancy Rate'],
                        name='Occupancy Rate (%)',
                        mode='lines+markers',
                        marker_color='#FFC107',
                        line=dict(width=3)
                    ),
                    secondary_y=True
                )
                
                # Set titles
                fig.update_layout(
                    title_text='Building Summary',
                    xaxis_title='Building'
                )
                
                fig.update_yaxes(title_text='Room Count', secondary_y=False)
                fig.update_yaxes(title_text='Occupancy Rate (%)', secondary_y=True)
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Export options
                if st.button("Generate Building Summary PDF"):
                    st.info("PDF export functionality would be implemented here (requires additional libraries)")
                    
                # Export as CSV
                building_csv = building_summary.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Building Summary CSV",
                    data=building_csv,
                    file_name=f"building_summary_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("No room data available for building summary")
