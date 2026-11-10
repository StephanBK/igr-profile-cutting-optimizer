import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import os
from datetime import datetime
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
import psycopg2
import psycopg2.extras
import hashlib

# Page config
st.set_page_config(
    page_title="IGR Profile Cutting Optimizer",
    page_icon="⚙️",
    layout="wide"
)

# Styling
st.markdown("""
<style>
    .stMetric {
        background-color: #495057 !important;
        border: 1px solid #6c757d !important;
        padding: 1rem !important;
        border-radius: 0.5rem !important;
    }
    .stMetric > div {
        background-color: #495057 !important;
        color: white !important;
    }
    .stMetric label {
        color: white !important;
    }
    .stMetric [data-testid="metric-container"] {
        background-color: #495057 !important;
        color: white !important;
    }
    .stMetric [data-testid="metric-container"] > div {
        color: white !important;
    }
    .stMultiSelect div[data-baseweb="tag"] {
        background-color: #28a745 !important;
        color: white !important;
    }
    .stButton > button {
        background-color: #28a745;
        color: white;
        border: none;
    }
    .profile-summary {
        background-color: #212529;
        color: #ffffff;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Constants
CUT_LOSS_90 = 3.0 / 16.0  # 3/16 inch for 90 degree cuts
CUT_LOSS_45 = 0.5  # 1/2 inch for 45 degree cuts

# Database login
# The login is a secret. It is NEVER written in this file.
# Set DATABASE_URL in the app's secret settings:
#   Streamlit Cloud: App settings > Secrets
#   Railway or local: environment variable DATABASE_URL
# Format: postgresql://USER:PASSWORD@HOST:5432/DBNAME
def get_database_url():
    try:
        if "DATABASE_URL" in st.secrets:
            return st.secrets["DATABASE_URL"]
    except Exception:
        pass  # No secrets file. Fall back to the environment variable.
    return os.environ.get("DATABASE_URL")


def connect_db():
    url = get_database_url()
    if not url:
        return None
    try:
        return psycopg2.connect(url, sslmode="require")
    except Exception:
        return None


def parse_excel(file):
    try:
        df = pd.read_excel(file, sheet_name='Cutfile', header=None)

        # Find header row
        header_row = None
        for i, row in df.iterrows():
            if pd.notna(row[2]) and str(row[2]).strip() == "Profile #":
                header_row = i
                break

        if header_row is None:
            st.error("Cannot find Profile # column")
            return None, None

        # Extract project info
        project_info = {}
        for i in range(header_row):
            if pd.notna(df.iloc[i, 0]):
                key = str(df.iloc[i, 0]).replace(":", "").strip()
                if len(df.columns) > 1 and pd.notna(df.iloc[i, 1]):
                    project_info[key] = df.iloc[i, 1]

        # Extract cutting data
        cutting_data = []
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            if pd.isna(row[0]) or row[0] == '':
                continue

            try:
                cutting_data.append({
                    'profile_num': str(row[2]).strip(),
                    'qty': int(float(row[5])),
                    'length_in': float(row[7]),
                    'cutting_angle': float(row[8]) if len(row) > 8 and pd.notna(row[8]) else 90.0,
                    'description': str(row[3]) if pd.notna(row[3]) else '',
                    'unit_type': str(row[9]).strip() if len(row) > 9 and pd.notna(row[9]) else '',
                    # FIXED: Unit type is column 9
                    'position': str(row[4]).strip() if len(row) > 4 and pd.notna(row[4]) else '',
                    # FIXED: Position is column 4
                })
            except:
                continue

        return project_info, cutting_data
    except Exception as e:
        st.error(f"Error parsing file: {e}")
        return None, None


def aggregate_profiles(cutting_data):
    profiles = {}

    for item in cutting_data:
        profile = item['profile_num']
        if profile not in profiles:
            profiles[profile] = []

        # Add cut loss based on angle
        cut_loss = CUT_LOSS_90 if item['cutting_angle'] == 90 else CUT_LOSS_45

        # Add each individual piece
        for _ in range(item['qty']):
            profiles[profile].append({
                'length': item['length_in'],
                'actual_length': item['length_in'] + cut_loss,
                'angle': item['cutting_angle'],
                'cut_loss': cut_loss,
                'description': item['description'],
                'unit_type': item['unit_type'],  # NEW: Pass through unit type
                'position': item['position']  # NEW: Pass through position
            })

    return profiles


def query_materials(conn, profile_numbers):
    if not conn:
        return {}

    materials = {}
    cursor = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    for profile in profile_numbers:
        query = """
        SELECT nickname, yield, quantity
        FROM materials 
        WHERE category = 1 AND nickname ILIKE %s
        ORDER BY yield DESC
        """
        cursor.execute(query, (f'%{profile}%',))
        results = cursor.fetchall()

        if results:
            materials[profile] = []
            for row in results:
                materials[profile].append({
                    'nickname': str(row['nickname']),
                    'yield_ft': float(row['yield']),
                    'yield_in': float(row['yield']) * 12,
                    'quantity': int(float(row['quantity']))
                })

    cursor.close()
    return materials


def optimize_cutting(pieces, selected_materials, available_materials, profile_num):
    if not pieces:
        return []

    # Sort pieces by size (largest first)
    pieces.sort(key=lambda x: x['actual_length'], reverse=True)

    # Determine stock lengths
    if 0 in selected_materials:  # "Other" selected
        if profile_num in available_materials and available_materials[profile_num]:
            # Use same yield as database options
            stock_length = max(m['yield_in'] for m in available_materials[profile_num])
        else:
            stock_length = 20 * 12  # Default 20ft
        stock_pieces = [stock_length] * 200  # Plenty of stock
    else:
        # Use selected database materials
        stock_pieces = []
        for idx in selected_materials:
            if idx > 0 and profile_num in available_materials:
                materials_list = available_materials[profile_num]
                if idx - 1 < len(materials_list):
                    material = materials_list[idx - 1]
                    for _ in range(material['quantity']):
                        stock_pieces.append(material['yield_in'])

        # Add buffer stock if needed
        if stock_pieces:
            longest = max(stock_pieces)
            stock_pieces.extend([longest] * 100)
        else:
            stock_pieces = [20 * 12] * 100  # Fallback

    # Cutting algorithm
    bins = []

    for piece in pieces:
        placed = False

        # Try existing bins
        for bin_data in bins:
            used = sum(p['actual_length'] for p in bin_data['pieces'])
            remaining = bin_data['stock_length'] - used
            if remaining >= piece['actual_length']:
                bin_data['pieces'].append(piece)
                placed = True
                break

        # Create new bin
        if not placed:
            suitable_stock = None
            for stock_len in stock_pieces:
                if stock_len >= piece['actual_length']:
                    suitable_stock = stock_len
                    break

            if suitable_stock is None:
                suitable_stock = max(stock_pieces)

            bins.append({
                'stock_length': suitable_stock,
                'pieces': [piece]
            })

    return bins


def create_diagram(bins, profile_num):
    if not bins:
        return None

    fig = make_subplots(
        rows=min(len(bins), 10),
        cols=1,
        subplot_titles=[f"Stock #{i + 1}: {b['stock_length']:.0f}\" ({b['stock_length'] / 12:.1f}ft)"
                        for i, b in enumerate(bins[:10])],
        vertical_spacing=0.08
    )

    colors = ['#28a745', '#17a2b8', '#fd7e14', '#e83e8c', '#6f42c1']

    for i, bin_data in enumerate(bins[:10]):
        row = i + 1

        for j, piece in enumerate(bin_data['pieces']):
            color = colors[j % len(colors)]

            # Main piece
            fig.add_trace(go.Bar(
                x=[piece['length']],
                y=[f"Stock {i + 1}"],
                orientation='h',
                marker_color=color,
                text=f"{piece['length']:.1f}\"",
                textposition='inside',
                showlegend=False,
                hovertemplate=f"Length: {piece['length']:.1f}\"<br>Angle: {piece['angle']:.0f}°<extra></extra>"
            ), row=row, col=1)

            # Cut loss
            fig.add_trace(go.Bar(
                x=[piece['cut_loss']],
                y=[f"Stock {i + 1}"],
                orientation='h',
                marker_color='#6c757d',
                text="CUT",
                textposition='inside',
                showlegend=False
            ), row=row, col=1)

        # Waste
        used = sum(p['actual_length'] for p in bin_data['pieces'])
        waste = bin_data['stock_length'] - used
        if waste > 0.1:
            fig.add_trace(go.Bar(
                x=[waste],
                y=[f"Stock {i + 1}"],
                orientation='h',
                marker_color='#adb5bd',
                text=f"Waste: {waste:.1f}\"",
                textposition='inside',
                showlegend=False
            ), row=row, col=1)

        fig.update_xaxes(
            range=[0, bin_data['stock_length'] * 1.05],
            title_text="Length (inches)" if i == min(len(bins), 10) - 1 else "",
            row=row, col=1
        )
        fig.update_yaxes(visible=False, row=row, col=1)

    fig.update_layout(
        height=80 * min(len(bins), 10) + 150,
        title_text=f"Profile {profile_num} Cutting Layout",
        barmode='stack',
        showlegend=False
    )

    return fig


def get_cutting_pattern_hash(pieces):
    """Generate a hash for a cutting pattern to identify identical patterns"""
    sorted_pieces = sorted(pieces, key=lambda x: (x['length'], x['angle']))
    pattern_str = ""
    for piece in sorted_pieces:
        pattern_str += f"{piece['length']:.3f}-{piece['angle']:.0f}"
    return hashlib.md5(pattern_str.encode()).hexdigest()


def create_pattern_diagram(pieces, stock_length):
    """Create a visual diagram for a cutting pattern using ReportLab graphics"""
    drawing = Drawing(500, 60)

    # Professional color scheme
    piece_colors = [
        colors.Color(1, 0.8, 0.8),  # Light red
        colors.Color(0.8, 0.8, 1),  # Light blue
        colors.Color(0.8, 1, 0.8),  # Light green
        colors.Color(1, 1, 0.8),  # Light yellow
        colors.Color(1, 0.8, 1),  # Light magenta
        colors.Color(0.8, 1, 1),  # Light cyan
        colors.Color(1, 0.9, 0.7),  # Light orange
        colors.Color(0.9, 0.8, 1)  # Light purple
    ]

    cut_color = colors.Color(1, 0.84, 0)  # Gold for cuts
    waste_color = colors.Color(0.83, 0.83, 0.83)  # Light gray for waste

    # Scale factor to fit drawing width
    scale = 480 / stock_length
    x_pos = 10
    y_pos = 10
    height = 30

    # Draw each piece and its cut loss
    for i, piece in enumerate(pieces):
        piece_color = piece_colors[i % len(piece_colors)]
        piece_width = piece['length'] * scale
        cut_width = piece['cut_loss'] * scale

        # Draw main piece
        rect = Rect(x_pos, y_pos, piece_width, height)
        rect.fillColor = piece_color
        rect.strokeColor = colors.black
        rect.strokeWidth = 0.5
        drawing.add(rect)

        # Add text on piece
        if piece_width > 30:
            text = String(x_pos + piece_width / 2, y_pos + height / 2 - 3,
                          f"{piece['length']:.1f}\"", textAnchor='middle')
            text.fontSize = 8
            text.fillColor = colors.black
            drawing.add(text)

        x_pos += piece_width

        # Draw cut loss
        cut_rect = Rect(x_pos, y_pos, cut_width, height)
        cut_rect.fillColor = cut_color
        cut_rect.strokeColor = colors.black
        cut_rect.strokeWidth = 0.5
        drawing.add(cut_rect)

        # Add cut text
        if cut_width > 15:
            cut_text = String(x_pos + cut_width / 2, y_pos + height / 2 - 3,
                              "CUT", textAnchor='middle')
            cut_text.fontSize = 6
            cut_text.fillColor = colors.black
            drawing.add(cut_text)

        x_pos += cut_width

    # Calculate and draw waste
    used_length = sum(p['actual_length'] for p in pieces)
    waste = stock_length - used_length

    if waste > 0.1:
        waste_width = waste * scale
        waste_rect = Rect(x_pos, y_pos, waste_width, height)
        waste_rect.fillColor = waste_color
        waste_rect.strokeColor = colors.black
        waste_rect.strokeWidth = 0.5
        drawing.add(waste_rect)

        # Add waste text
        if waste_width > 30:
            waste_text = String(x_pos + waste_width / 2, y_pos + height / 2 - 3,
                                f"Waste: {waste:.1f}\"", textAnchor='middle')
            waste_text.fontSize = 7
            waste_text.fillColor = colors.black
            drawing.add(waste_text)

    return drawing


def create_labels_pdf(all_results, project_info):
    """Create PDF labels for Avery 5160 format with enhanced information"""
    buffer = io.BytesIO()

    # Avery 5160 specifications (3 columns, 10 rows per page)
    page_width, page_height = letter
    label_width = 2.625 * inch
    label_height = 1.0 * inch
    left_margin = 0.1875 * inch
    top_margin = 0.5 * inch
    h_spacing = 0.125 * inch
    v_spacing = 0.0 * inch

    c = canvas.Canvas(buffer, pagesize=letter)

    labels_data = []
    # Use Project Name from the header info that was read from the Excel file - NO DEFAULTS
    project_name = project_info.get('Project Name', '')

    # Collect all pieces for labels with enhanced information
    for profile_num, result in all_results.items():
        bins = result['bins']
        for bin_data in bins:
            for piece in bin_data['pieces']:
                labels_data.append({
                    'project': project_name,
                    'profile': profile_num,
                    'length': f"{piece['length']:.3f}\"",
                    'unit': piece.get('unit_type', 'Unknown'),  # NEW: Use unit_type from input file
                    'position': piece.get('position', 'N/A')  # NEW: Use position from input file
                })

    # Generate labels
    labels_per_page = 30  # 3 cols x 10 rows
    label_count = 0

    for i, label in enumerate(labels_data):
        if i > 0 and i % labels_per_page == 0:
            c.showPage()

        # Calculate position
        row = (i % labels_per_page) // 3
        col = (i % labels_per_page) % 3

        x = left_margin + col * (label_width + h_spacing)
        y = page_height - top_margin - (row + 1) * (label_height + v_spacing)

        # Draw label border (optional)
        c.setStrokeColor(colors.lightgrey)
        c.rect(x, y, label_width, label_height)

        # Add text with enhanced layout - smaller fonts and tighter spacing
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 7)  # FIXED: Even smaller font for project name
        c.drawString(x + 5, y + label_height - 12, project_name)  # FIXED: Tighter spacing

        c.setFont("Helvetica", 7)  # FIXED: Smaller font for all other text
        c.drawString(x + 5, y + label_height - 24, f"Profile: {label['profile']}")  # FIXED: Tighter spacing
        c.drawString(x + 5, y + label_height - 36, f"Length: {label['length']}")  # FIXED: Tighter spacing
        c.drawString(x + 5, y + label_height - 48, f"Unit: {label['unit']}")  # FIXED: Tighter spacing
        c.drawString(x + 5, y + label_height - 60, f"Position: {label['position']}")  # FIXED: Tighter spacing

        label_count += 1

    c.save()
    buffer.seek(0)
    return buffer, label_count


def create_excel_report(all_results, project_info, available_materials):
    """Create Excel report with cutting patterns in table format"""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Create formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#366092',
            'font_color': 'white',
            'border': 1
        })

        cell_format = workbook.add_format({
            'border': 1,
            'align': 'center'
        })

        waste_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'bg_color': '#FFCCCC'
        })

        shortage_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FF6B6B',
            'font_color': 'white',
            'border': 1,
            'align': 'center'
        })

        # Create summary data
        summary_data = []

        for profile_num, result in all_results.items():
            bins = result['bins']
            selected_materials = result['selected_materials']
            stock_num = 1

            # Calculate shortage for this profile
            total_available = 0
            shortage = 0
            if 0 not in selected_materials:  # Not "Other" option
                for idx in selected_materials:
                    if idx > 0 and profile_num in available_materials:
                        materials_list = available_materials[profile_num]
                        if idx - 1 < len(materials_list):
                            total_available += materials_list[idx - 1]['quantity']

                if len(bins) > total_available:
                    shortage = len(bins) - total_available

            for bin_data in bins:
                stock_length_ft = bin_data['stock_length'] / 12

                # Add each piece in this stock
                for i, piece in enumerate(bin_data['pieces']):
                    summary_data.append({
                        'Profile': profile_num,
                        'Stock #': f"{profile_num}-{stock_num}",
                        'Stock Length (ft)': f"{stock_length_ft:.1f}",
                        'Stock Length (in)': f"{bin_data['stock_length']:.1f}",
                        'Cut #': i + 1,
                        'Piece Length (in)': f"{piece['length']:.3f}",
                        'Cut Angle': f"{piece['angle']:.0f}°",
                        'Cut Loss (in)': f"{piece['cut_loss']:.3f}",
                        'Type': 'PIECE',
                        'Shortage': f"{shortage} additional needed" if shortage > 0 and stock_num == 1 else ''
                    })

                # Add waste entry
                used_length = sum(p['actual_length'] for p in bin_data['pieces'])
                waste = bin_data['stock_length'] - used_length

                if waste > 0.01:
                    summary_data.append({
                        'Profile': profile_num,
                        'Stock #': f"{profile_num}-{stock_num}",
                        'Stock Length (ft)': f"{stock_length_ft:.1f}",
                        'Stock Length (in)': f"{bin_data['stock_length']:.1f}",
                        'Cut #': '',
                        'Piece Length (in)': f"{waste:.3f}",
                        'Cut Angle': '',
                        'Cut Loss (in)': '',
                        'Type': 'WASTE',
                        'Shortage': ''
                    })

                stock_num += 1

        # Create DataFrame and write to Excel
        df = pd.DataFrame(summary_data)
        df.to_excel(writer, sheet_name='Cutting Report', index=False)

        # Format the worksheet
        worksheet = writer.sheets['Cutting Report']

        # Write headers with formatting
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Format data rows
        for row in range(1, len(df) + 1):
            for col in range(len(df.columns)):
                if df.iloc[row - 1]['Type'] == 'WASTE':
                    worksheet.write(row, col, df.iloc[row - 1, col], waste_format)
                elif df.iloc[row - 1]['Shortage'] and 'additional needed' in str(df.iloc[row - 1]['Shortage']):
                    worksheet.write(row, col, df.iloc[row - 1, col], shortage_format)
                else:
                    worksheet.write(row, col, df.iloc[row - 1, col], cell_format)

        # Auto-adjust column widths
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, max(len(col), 12))

    buffer.seek(0)
    return buffer


def create_cutting_patterns_pdf(all_results, project_info, available_materials):
    """Create PDF showing cutting patterns for all profiles"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5 * inch, bottomMargin=0.5 * inch)
    story = []
    styles = getSampleStyleSheet()

    # Title
    title_style = styles['Title']
    title_style.alignment = TA_CENTER
    project_name = project_info.get('Project Name', '')
    story.append(
        Paragraph(f"IGR Cutting Patterns - {project_name}" if project_name else "IGR Cutting Patterns", title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Paragraph(f"Cut Loss: 90°={CUT_LOSS_90:.3f}\", 45°={CUT_LOSS_45:.1f}\"", styles['Normal']))
    story.append(Spacer(1, 12))

    # Process each profile
    for profile_num, result in all_results.items():
        bins = result['bins']
        selected_materials = result['selected_materials']
        if not bins:
            continue

        # Calculate shortage for this profile
        total_available = 0
        shortage = 0
        if 0 not in selected_materials:  # Not "Other" option
            for idx in selected_materials:
                if idx > 0 and profile_num in available_materials:
                    materials_list = available_materials[profile_num]
                    if idx - 1 < len(materials_list):
                        total_available += materials_list[idx - 1]['quantity']

            if len(bins) > total_available:
                shortage = len(bins) - total_available

        # Profile header
        story.append(Paragraph(f"Profile {profile_num} Cutting Patterns", styles['Heading1']))

        # Add shortage warning if needed
        if shortage > 0:
            shortage_style = styles['Normal']
            shortage_style.textColor = colors.red
            story.append(Paragraph(f"⚠️ WARNING: Need {shortage} additional stock pieces", shortage_style))

        story.append(Spacer(1, 12))

        # Group identical cutting patterns
        pattern_groups = {}
        for bin_data in bins:
            pattern_hash = get_cutting_pattern_hash(bin_data['pieces'])
            if pattern_hash not in pattern_groups:
                pattern_groups[pattern_hash] = {
                    'stock_length': bin_data['stock_length'],
                    'pieces': bin_data['pieces'],
                    'count': 1
                }
            else:
                pattern_groups[pattern_hash]['count'] += 1

        # Sort by frequency (most used first)
        sorted_patterns = dict(sorted(pattern_groups.items(),
                                      key=lambda x: x[1]['count'],
                                      reverse=True))

        pattern_num = 1
        for pattern_hash, pattern_data in sorted_patterns.items():
            # Pattern header
            stock_length_ft = pattern_data['stock_length'] / 12
            header_text = f"Pattern #{pattern_num} - {stock_length_ft:.1f}ft ({pattern_data['stock_length']:.1f}\") Stock - Used {pattern_data['count']} times"

            pattern_content = []
            pattern_content.append(Paragraph(header_text, styles['Heading2']))

            # Create visual diagram
            diagram = create_pattern_diagram(pattern_data['pieces'], pattern_data['stock_length'])
            pattern_content.append(diagram)
            pattern_content.append(Spacer(1, 8))

            # Create detailed table
            table_data = []
            table_data.append(['Cut #', 'Length (in)', 'Angle', 'Cut Loss', 'Running Total'])

            running_total = 0
            for i, piece in enumerate(pattern_data['pieces']):
                running_total += piece['actual_length']
                table_data.append([
                    str(i + 1),
                    f"{piece['length']:.3f}\"",
                    f"{piece['angle']:.0f}°",
                    f"{piece['cut_loss']:.3f}\"",
                    f"{running_total:.3f}\""
                ])

            # Add waste row
            waste = pattern_data['stock_length'] - running_total
            table_data.append(['', 'WASTE', '', f"{waste:.3f}\"", f"{pattern_data['stock_length']:.3f}\""])

            # Create table
            table = Table(table_data, colWidths=[0.6 * inch, 1.2 * inch, 0.8 * inch, 0.8 * inch, 1.1 * inch])

            # Style the table with bright colors
            bright_colors = [
                colors.Color(1, 0.8, 0.8),  # Light red
                colors.Color(0.8, 0.8, 1),  # Light blue
                colors.Color(0.8, 1, 0.8),  # Light green
                colors.Color(1, 1, 0.8),  # Light yellow
                colors.Color(1, 0.8, 1),  # Light magenta
                colors.Color(0.8, 1, 1),  # Light cyan
                colors.Color(1, 0.9, 0.7),  # Light orange
                colors.Color(0.9, 0.8, 1)  # Light purple
            ]

            table_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.2, 0.8)),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('BACKGROUND', (0, 1), (-1, -2), colors.Color(0.98, 0.98, 0.95)),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
            ]

            # Color each piece row differently
            for i in range(1, len(table_data) - 1):
                color_idx = (i - 1) % len(bright_colors)
                table_style.append(('BACKGROUND', (0, i), (-1, i), bright_colors[color_idx]))

            # Special styling for waste row
            table_style.append(('BACKGROUND', (0, -1), (-1, -1), colors.Color(0.9, 0.9, 0.9)))
            table_style.append(('FONTNAME', (0, -1), (1, -1), 'Helvetica-Bold'))

            table.setStyle(TableStyle(table_style))
            pattern_content.append(table)

            # Add efficiency info
            material_used = sum(p['actual_length'] for p in pattern_data['pieces'])
            efficiency = (material_used / pattern_data['stock_length']) * 100
            efficiency_text = f"Efficiency: {efficiency:.1f}% | Waste: {waste:.3f}\" ({waste / 12:.3f}ft) | Pieces: {len(pattern_data['pieces'])}"
            pattern_content.append(Spacer(1, 6))
            pattern_content.append(Paragraph(efficiency_text, styles['Normal']))
            pattern_content.append(Spacer(1, 12))

            # Keep pattern together on page
            story.append(KeepTogether(pattern_content))

            pattern_num += 1

        # Add space between profiles
        story.append(Spacer(1, 24))

    # Build PDF
    doc.build(story)
    buffer.seek(0)
    return buffer


# Initialize session state
if 'results' not in st.session_state:
    st.session_state.results = None

# Main App
st.title("⚙️ IGR Profile Cutting Optimizer")
st.markdown("**Professional aluminum profile cutting optimization with database integration**")

# Connect to database
conn = connect_db()
if conn:
    st.sidebar.success("Database Connected")
else:
    st.sidebar.error("Database Offline")

# Template download button in sidebar
st.sidebar.divider()
st.sidebar.subheader("Template")
try:
    with open("IGR_template.xlsx", "rb") as template_file:
        st.sidebar.download_button(
            label="📥 Download Template",
            data=template_file.read(),
            file_name="IGR_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except FileNotFoundError:
    st.sidebar.error("Template file not found")

# File upload
uploaded_file = st.file_uploader("Upload IGR Cut File", type=['xlsx'])

if uploaded_file:
    project_info, cutting_data = parse_excel(uploaded_file)

    if project_info and cutting_data:
        st.success(f"File loaded: {len(cutting_data)} items")

        # Aggregate by profile
        profiles = aggregate_profiles(cutting_data)

        # Profile summary
        st.subheader("Profile Summary")
        summary_html = '<div class="profile-summary">'
        for profile_num, pieces in profiles.items():
            angle_90 = sum(1 for p in pieces if p['angle'] == 90)
            angle_45 = sum(1 for p in pieces if p['angle'] == 45)
            summary_html += f'<div>Profile {profile_num}: {len(pieces)} pieces (90°: {angle_90}, 45°: {angle_45})</div>'
        summary_html += '</div>'
        st.markdown(summary_html, unsafe_allow_html=True)

        # Query materials
        if conn:
            available_materials = query_materials(conn, list(profiles.keys()))

            if available_materials:
                st.success(f"Materials found for {len(available_materials)} profiles")

                # Material selection
                st.subheader("Material Selection")
                selections = {}

                for profile_num in profiles.keys():
                    if profile_num in available_materials:
                        materials = available_materials[profile_num]

                        options = ["Other (Same lengths as database options)"]
                        for mat in materials:
                            options.append(f"{mat['nickname']} - {mat['yield_ft']:.0f}ft - Qty: {mat['quantity']}")

                        selected = st.multiselect(
                            f"Profile {profile_num}",
                            range(len(options)),
                            format_func=lambda x: options[x],
                            key=f"sel_{profile_num}"
                        )
                        selections[profile_num] = selected
                    else:
                        st.warning(f"No materials for Profile {profile_num}")
                        selections[profile_num] = []

                # Optimize button
                if st.button("Optimize Cutting", type="primary"):
                    results = {}

                    with st.spinner("Optimizing profiles..."):
                        for profile_num, pieces in profiles.items():
                            selected_materials = selections.get(profile_num, [])
                            bins = optimize_cutting(pieces, selected_materials, available_materials, profile_num)

                            total_stock = sum(b['stock_length'] for b in bins)
                            total_material = sum(p['actual_length'] for b in bins for p in b['pieces'])
                            efficiency = (total_material / total_stock * 100) if total_stock > 0 else 0

                            results[profile_num] = {
                                'bins': bins,
                                'pieces': pieces,
                                'efficiency': efficiency,
                                'selected_materials': selected_materials
                            }

                    # Store results
                    st.session_state.results = results
                    st.session_state.project_info = project_info
                    st.session_state.available_materials = available_materials

# Display results if they exist
if st.session_state.results:
    st.header("Optimization Results")
    results = st.session_state.results

    # Overall summary
    total_profiles = len(results)
    total_pieces = sum(len(r['pieces']) for r in results.values())
    total_stock_pieces = sum(len(r['bins']) for r in results.values())
    avg_efficiency = sum(r['efficiency'] for r in results.values()) / total_profiles if total_profiles > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Profiles", total_profiles)
    with col2:
        st.metric("Total Pieces", total_pieces)
    with col3:
        st.metric("Stock Pieces", total_stock_pieces)
    with col4:
        st.metric("Avg Efficiency", f"{avg_efficiency:.1f}%")

    # Profile-specific results
    st.subheader("Results by Profile")

    tabs = st.tabs([f"Profile {p}" for p in results.keys()])

    for tab, (profile_num, result) in zip(tabs, results.items()):
        with tab:
            bins = result['bins']
            pieces = result['pieces']
            efficiency = result['efficiency']
            selected_materials = result['selected_materials']

            if bins:
                # Calculate stats
                total_stock = sum(b['stock_length'] for b in bins)
                total_material = sum(p['actual_length'] for b in bins for p in b['pieces'])
                waste = total_stock - total_material

                # Display metrics
                col_p1, col_p2, col_p3, col_p4 = st.columns(4)
                with col_p1:
                    st.metric("Pieces Cut", len(pieces))
                with col_p2:
                    st.metric("Efficiency", f"{efficiency:.1f}%")
                with col_p3:
                    st.metric("Stock Used", len(bins))
                with col_p4:
                    st.metric("Total Waste", f"{waste:.0f}\"")

                # Show ordering info
                if 0 in selected_materials:  # Other selected
                    stock_ft = bins[0]['stock_length'] / 12
                    st.info(f"Order {len(bins)} pieces of {stock_ft:.0f}ft stock")
                else:
                    # Calculate shortage
                    total_available = 0
                    for idx in selected_materials:
                        if idx > 0 and profile_num in st.session_state.available_materials:
                            materials_list = st.session_state.available_materials[profile_num]
                            if idx - 1 < len(materials_list):
                                total_available += materials_list[idx - 1]['quantity']

                    if len(bins) > total_available:
                        shortage = len(bins) - total_available
                        st.warning(f"Need {shortage} additional stock pieces")
                    else:
                        st.success("Sufficient stock available")

                # Cutting diagram
                fig = create_diagram(bins, profile_num)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.error(f"No solution found for Profile {profile_num}")

    # Export section
    st.divider()
    st.subheader("Export Results")

    col_e1, col_e2, col_e3 = st.columns(3)

    with col_e1:
        if st.button("Generate Labels"):
            try:
                labels_buffer, label_count = create_labels_pdf(st.session_state.results, st.session_state.project_info)
                st.download_button(
                    "Download Labels PDF",
                    labels_buffer.getvalue(),
                    f"IGR_labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    "application/pdf"
                )
                st.success(f"Generated {label_count} labels")
            except Exception as e:
                st.error(f"Error generating labels: {str(e)}")

    with col_e2:
        if st.button("Generate Excel"):
            try:
                excel_buffer = create_excel_report(st.session_state.results, st.session_state.project_info,
                                                   st.session_state.available_materials)
                st.download_button(
                    "Download Excel Report",
                    excel_buffer.getvalue(),
                    f"IGR_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Excel report generated!")
            except Exception as e:
                st.error(f"Error generating Excel: {str(e)}")

    with col_e3:
        if st.button("Generate PDF Patterns"):
            try:
                pdf_buffer = create_cutting_patterns_pdf(st.session_state.results, st.session_state.project_info,
                                                         st.session_state.available_materials)
                st.download_button(
                    "Download PDF Patterns",
                    pdf_buffer.getvalue(),
                    f"IGR_patterns_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    "application/pdf"
                )
                st.success("PDF patterns generated!")
            except Exception as e:
                st.error(f"Error generating PDF: {str(e)}")