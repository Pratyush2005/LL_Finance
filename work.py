import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import os
import requests
from urllib.parse import urlparse


# --- Configuration & Color Palette ---
NAVY_BLUE = '#001f3f'
GREEN = '#2ECC40'
RED = '#FF4136'
AMBER = '#FF851B'
GRAY = '#AAAAAA'
LIGHT_GRAY = '#F0F0F0'

def first_non_nan_url(*values):
    for v in values:
        if isinstance(v, str) and v.strip().lower().startswith(('http://', 'https://')):
            return v.strip()
    return None

def download_image(url, dest_path):
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(r.content)
        return dest_path
    except Exception:
        return None

# --- Data Derivation Functions (for Hyper-Personalization) ---
def get_company_metrics(employees, industry):
    """Derives a suite of personalized metrics based on company size and industry."""
    metrics = {}
    industry = str(industry).lower()

    # 1. Processing Time (Days)
    if employees < 50:
        metrics['processing_time'] = 21
    elif 50 <= employees < 250:
        metrics['processing_time'] = 15
    else:
        metrics['processing_time'] = 10
    
    # 2. Cost Per Invoice ($)
    metrics['cost_per_invoice'] = metrics['processing_time'] * 1.8 + 5 # Simplified correlation
    if 'financial' in industry:
        metrics['cost_per_invoice'] *= 1.2

    # 3. First-Time Match Rate (%)
    if 'manufacturing' in industry:
        metrics['first_time_match_rate'] = 35
    elif 'tech' in industry:
        metrics['first_time_match_rate'] = 65
    else:
        metrics['first_time_match_rate'] = 50

    # 4. AP Efficiency Score (%) - A weighted score
    score = ( (5 / metrics['processing_time']) * 40 + \
              (12 / metrics['cost_per_invoice']) * 40 + \
              (metrics['first_time_match_rate'] / 85) * 20 )
    metrics['efficiency_score'] = min(int(score), 95)

    # 5. Potential Annual Savings ($)
    invoices_per_employee_per_month = 20
    total_invoices = employees * invoices_per_employee_per_month * 12
    savings_per_invoice = metrics['cost_per_invoice'] - 5 # Assuming optimized cost is $5
    metrics['annual_savings'] = int(total_invoices * savings_per_invoice)

    return metrics

# --- Chart Generation Functions ---
# EDITED FUNCTION
def create_donut_chart(value, benchmark, title, color, filename):
    """Creates and saves a donut chart."""
    fig, ax = plt.subplots(figsize=(2, 2))
    
    # FIX: Use max(0, ...) to prevent the second wedge size from being negative.
    # This handles cases where the value is greater than the benchmark.
    sizes = [value, max(0, benchmark - value)]
    
    main_color = color
    remaining_color = LIGHT_GRAY
    
    wedges, texts, autotexts = ax.pie(sizes, labels=['', ''], autopct='%1.0f%%', startangle=90, 
                                      colors=[main_color, remaining_color],
                                      wedgeprops=dict(width=0.3, edgecolor='w'),
                                      textprops={'color': NAVY_BLUE, 'fontsize': 14, 'weight': 'bold'})
    
    # Remove percentage from the smaller wedge if it exists
    if len(autotexts) > 1:
        autotexts[1].set_text('')

    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig.gca().add_artist(centre_circle)
    
    # Display the actual value in the center
    ax.text(0, 0, f'{int(value)}', ha='center', va='center', fontsize=20, weight='bold', color=NAVY_BLUE)

    ax.axis('equal')
    plt.title(title, fontsize=10, color=NAVY_BLUE)
    plt.savefig(filename, transparent=True, bbox_inches='tight')
    plt.close()

def create_peer_comparison_chart(company_score, filename):
    """Creates and saves a horizontal bar chart for peer comparison."""
    peers = {'Your Company': company_score, 'Competitor A': 81, 'Industry Leader': 95}
    names = list(peers.keys())
    scores = list(peers.values())
    fig, ax = plt.subplots(figsize=(6, 1.5))
    bars = ax.barh(names, scores, color=[NAVY_BLUE, GRAY, GREEN])
    ax.invert_yaxis()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.xaxis.set_ticks([])
    ax.tick_params(axis='y', length=0, labelsize=10, colors=NAVY_BLUE)
    ax.set_xlim(0, 100)
    for bar in bars:
        width = bar.get_width()
        ax.text(width + 3, bar.get_y() + bar.get_height()/2, f'{int(width)}%', 
                ha='left', va='center', color=NAVY_BLUE, weight='bold')
    plt.savefig(filename, transparent=True, bbox_inches='tight')
    plt.close()

def create_monthly_savings_visual(filename, current_monthly, optimized_monthly):
    savings = max(0, current_monthly - optimized_monthly)
    fig, ax = plt.subplots(figsize=(3, 1.8))
    ax.bar(['Current'], [current_monthly], color=RED)
    ax.bar(['Optimized'], [optimized_monthly], color=GREEN)
    ax.set_ylabel('Monthly Cost ($)', fontsize=10, color=NAVY_BLUE)
    ax.set_title('Monthly Savings', fontsize=12, color=NAVY_BLUE)
    for spine in ax.spines.values():
        spine.set_visible(False)
    # Big savings text
    ax.text(0.5, 0.9, f"Savings: ${int(savings):,}", ha='center', va='center', transform=ax.transAxes, fontsize=14, weight='bold', color=GREEN)
    plt.tight_layout()
    plt.savefig(filename, transparent=True, bbox_inches='tight')
    plt.close()

def create_savings_thermometer(filename, current, optimized):
    """Creates a thermometer-style bar comparing current vs optimized cost."""
    fig, ax = plt.subplots(figsize=(1.5, 4))
    ax.barh(0, current, color=RED, height=0.6, label='Current')
    ax.barh(0, optimized, color=GREEN, height=0.6, label='Optimized')
    ax.set_xlim(0, max(current, optimized)*1.2)
    ax.set_yticks([])
    ax.set_xlabel('Cost Per Invoice ($)', fontsize=10, color=NAVY_BLUE)
    ax.legend(loc='upper right', fontsize=8)
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, transparent=True, bbox_inches='tight')
    plt.close()

def create_money_leak_visualization(filename, brand_green=GREEN, brand_red=RED):
    """Creates a funnel-style visualization showing leaking vs optimized funnel."""
    stages = ['Invoices Received', 'Processed', 'Matched', 'Paid']
    values = [100, 80, 60, 55]  # Example values, showing leakages
    optimized = [100, 95, 90, 85]

    fig, ax = plt.subplots(figsize=(3, 3))
    ax.invert_yaxis()
    ax.barh(stages, optimized, color=brand_green, alpha=0.3, label='Optimized')
    ax.barh(stages, values, color=brand_red, label='Current')
    ax.set_xlabel('Percentage (%)', fontsize=10, color=NAVY_BLUE)
    ax.set_xlim(0, 110)
    ax.legend(fontsize=8)
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, transparent=True, bbox_inches='tight')
    plt.close()

# --- PDF Generation Class ---
class PDFReport(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.add_font('DejaVu', '', 'DejaVuSans.ttf')
        self.add_font('DejaVu', 'B', 'DejaVuSans-Bold.ttf')
        self.add_font('DejaVu', 'I', 'DejaVuSans-Oblique.ttf')

    def header(self):
        pass

    def footer(self):
        self.set_y(-15)
        self.set_font('DejaVu', 'I', 8)
        self.set_text_color(128)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')
        self.set_x(-60)
        self.cell(0, 10, f'Based on Q4 2024 data', align='R')

    def create_page_1(self, company_name, metrics, chart_files, logo_path=None, urgent=False):
        self.add_page()
        # --- Top section: logo and efficiency donut ---
        # Logo at (15, 15), w=30
        if logo_path and os.path.isfile(logo_path):
            self.image(logo_path, x=15, y=15, w=30)
        # Efficiency donut at (165, 15), w=35
        if chart_files.get('efficiency_meter') and os.path.isfile(chart_files['efficiency_meter']):
            self.image(chart_files['efficiency_meter'], x=165, y=15, w=35)

        # Headline below logo/donut at y=60
        self.set_xy(0, 60)
        self.set_font('DejaVu', 'B', 22)
        self.set_text_color(0, 31, 63)
        headline = f"{company_name}’s AP Efficiency Score: {metrics['efficiency_score']}%"
        self.cell(210, 11, headline, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        # Subheadline at y=75
        self.set_xy(0, 75)
        self.set_font('DejaVu', '', 12)
        self.set_text_color(128)
        subheadline = f"Industry Average: 78% | Potential Annual Savings: ${metrics['annual_savings']:,}"
        self.cell(210, 7, subheadline, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        # URGENT stamp if < 50% efficiency
        if urgent:
            self.set_draw_color(255, 65, 54)
            self.set_text_color(255, 65, 54)
            self.set_line_width(1)
            self.set_xy(140, 60)
            self.cell(50, 12, 'URGENT', border=1, align='C')
            self.set_text_color(0, 31, 63)

        # --- Middle visuals: thermometer and funnel side by side ---
        # Thermometer at (25, 95), w=70
        if chart_files.get('thermometer') and os.path.isfile(chart_files['thermometer']):
            self.image(chart_files['thermometer'], x=25, y=95, w=70)
        # Funnel at (115, 95), w=70
        if chart_files.get('funnel') and os.path.isfile(chart_files['funnel']):
            self.image(chart_files['funnel'], x=115, y=95, w=70)

        # --- Bottom donuts in a row with fixed y ---
        # Each donut: width 50, y=180; x = 25, 90, 155
        if chart_files.get('cost') and os.path.isfile(chart_files['cost']):
            self.image(chart_files['cost'], x=25, y=180, w=50)
        if chart_files.get('time') and os.path.isfile(chart_files['time']):
            self.image(chart_files['time'], x=90, y=180, w=50)
        if chart_files.get('match') and os.path.isfile(chart_files['match']):
            self.image(chart_files['match'], x=155, y=180, w=50)

        # Captions directly under each donut at y=225
        self.set_font('DejaVu', '', 10)
        self.set_text_color(128)
        self.set_xy(25, 225)
        self.cell(50, 6, "vs Industry Average ($12)", align='C')
        self.set_xy(90, 225)
        self.cell(50, 6, "vs Benchmark (5 days)", align='C')
        self.set_xy(155, 225)
        self.cell(50, 6, "vs Best Practice (85%)", align='C')

        # --- Peer comparison and monthly savings visuals ---
        # Peer comparison at (25, 240), w=80
        if chart_files.get('peer') and os.path.isfile(chart_files['peer']):
            self.image(chart_files['peer'], x=25, y=240, w=80)
        # Monthly savings at (125, 240), w=80
        if chart_files.get('monthly_savings') and os.path.isfile(chart_files['monthly_savings']):
            self.image(chart_files['monthly_savings'], x=125, y=240, w=80)

        # Add whitespace at bottom, push content away from footer
        self.set_y(270)
        self.ln(10)

    def create_page_2(self):
        self.add_page()
        # Title centered at top
        self.set_font('DejaVu', 'B', 20)
        self.set_text_color(0, 31, 63)
        self.cell(0, 15, "The ‘Here’s How’ Roadmap", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

        # Better spacing: process flow at y=40, bigger boxes and more spacing
        steps = [
            {'name': 'Receipt', 'color': (255, 204, 204), 'time': '3 days', 'cost': '$4'},
            {'name': 'Data Entry', 'color': (255, 204, 204), 'time': '4 days', 'cost': '$5'},
            {'name': 'Approval', 'color': (255, 236, 179), 'time': '5 days', 'cost': '$3'},
            {'name': 'Payment', 'color': (204, 255, 204), 'time': '2 days', 'cost': '$2'},
            {'name': 'Filing', 'color': (204, 255, 204), 'time': '1 day', 'cost': '$1'}
        ]
        box_width = 38
        box_height = 22
        start_x = 15
        start_y = 40
        self.set_font('DejaVu', 'B', 14)
        for i, step in enumerate(steps):
            self.set_fill_color(*step['color'])
            self.set_xy(start_x + i*box_width, start_y)
            self.cell(box_width, box_height, step['name'], border=1, align='C', fill=1)
        # Draw arrows between boxes
        arrow_y = start_y + box_height/2
        for i in range(len(steps)-1):
            x1 = start_x + (i+1)*box_width - 3
            y1 = arrow_y
            self.set_draw_color(0, 31, 63)
            self.line(x1, y1, x1 + 8, y1)
            # Draw arrowhead
            self.line(x1 + 8, y1, x1 + 4, y1 - 4)
            self.line(x1 + 8, y1, x1 + 4, y1 + 4)
        # Time/cost annotations under each step
        self.set_font('DejaVu', '', 10)
        for i, step in enumerate(steps):
            self.set_xy(start_x + i*box_width, start_y + box_height)
            self.cell(box_width, 6, f"Time: {step['time']}", border=0, align='C')
            self.set_xy(start_x + i*box_width, start_y + box_height + 6)
            self.cell(box_width, 6, f"Cost: {step['cost']}", border=0, align='C')

        # Add vertical spacing before quick wins to avoid overlap
        self.set_y(start_y + box_height + 18)
        self.ln(20)

        # Quick wins section: bullet points with left alignment, indent at x=20, line height=8
        self.set_font('DejaVu', 'B', 14)
        self.cell(0, 10, "3 Changes You Can Make This Week", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_font('DejaVu', '', 12)
        quick_wins = [
            ("Email Parsing:", "Centralize Invoices: Set up invoice@company.com to save 3+ hours daily."),
            ("Approval Matrix:", "Create an Approval Matrix: Cut approval time by 60% with a simple template."),
            ("Exception Tracking:", "Track Exceptions: Reduce errors by 40% with a basic exception tracker."),
        ]
        line_y = self.get_y()
        for label, text in quick_wins:
            self.set_xy(20, line_y)
            self.set_font('DejaVu', 'B', 12)
            self.cell(32, 8, label, align='L', border=0)
            self.set_font('DejaVu', '', 12)
            self.cell(0, 8, f" {text}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            line_y += 8

# --- Main Execution ---
def process_data_and_generate_reports(input_file_path):
    try:
        print("Looking for file at:", input_file_path)
        print("Exists?", os.path.exists(input_file_path))
        df = pd.read_excel(input_file_path)
        os.makedirs('reports', exist_ok=True)
        os.makedirs('img', exist_ok=True)
        pdf_filenames = []
        for index, row in df.iterrows():
            company_name = row.get('name', f'Company_{index}')
            employees = row.get('organization/estimated_num_employees', 50)
            industry = row.get('organization/industry', 'General')
            if pd.isna(employees): employees = 50
            metrics = get_company_metrics(employees, industry)
            safe_name = re.sub(r'[\\/*?:"<>|]', "", company_name)

            # Try to pull brand colors from sheet if available
            brand_primary = row.get('brand_primary', NAVY_BLUE)
            brand_secondary = row.get('brand_secondary', GREEN)
            if not isinstance(brand_primary, str) or not brand_primary.startswith('#'):
                brand_primary = NAVY_BLUE
            if not isinstance(brand_secondary, str) or not brand_secondary.startswith('#'):
                brand_secondary = GREEN

            # Logo URL from any plausible column
            logo_url = first_non_nan_url(row.get('logo'), row.get('logo_url'), row.get('organization/logo') if 'organization/logo' in df.columns else None, row.get('image'), row.get('image_url'))
            logo_path = None
            if logo_url:
                os.makedirs('img', exist_ok=True)
                candidate_logo = f"img/{safe_name}_logo.png"
                downloaded = download_image(logo_url, candidate_logo)
                if downloaded:
                    logo_path = downloaded

            chart_files = {
                'cost': f"img/{safe_name}_cost.png",
                'time': f"img/{safe_name}_time.png",
                'match': f"img/{safe_name}_match.png",
                'peer': f"img/{safe_name}_peer.png",
                'thermometer': f"img/{safe_name}_thermo.png",
                'funnel': f"img/{safe_name}_funnel.png",
                'efficiency_meter': f"img/{safe_name}_efficiency_meter.png",
                'monthly_savings': f"img/{safe_name}_monthly_savings.png",
            }

            # Donut charts (respect the requested colors)
            create_donut_chart(int(metrics['cost_per_invoice']), 12, 'Cost Per Invoice', RED, chart_files['cost'])
            create_donut_chart(metrics['processing_time'], 5, 'Processing Time (days)', AMBER, chart_files['time'])
            create_donut_chart(metrics['first_time_match_rate'], 85, 'First-Time Match Rate', RED, chart_files['match'])
            create_peer_comparison_chart(metrics['efficiency_score'], chart_files['peer'])

            # Savings thermometer & money leak using brand colors
            create_savings_thermometer(chart_files['thermometer'], metrics['cost_per_invoice'], 5)
            create_money_leak_visualization(chart_files['funnel'], brand_green=brand_secondary, brand_red=RED)

            # Efficiency meter donut chart for overlay
            create_donut_chart(metrics['efficiency_score'], 100, 'Efficiency Score', GREEN, chart_files['efficiency_meter'])

            # Monthly savings calculator visual
            invoices_per_employee_per_month = 20
            current_monthly = int(employees * invoices_per_employee_per_month * metrics['cost_per_invoice'])
            optimized_monthly = int(employees * invoices_per_employee_per_month * 5)
            create_monthly_savings_visual(chart_files['monthly_savings'], current_monthly, optimized_monthly)

            pdf = PDFReport()
            urgent_flag = metrics['efficiency_score'] < 50 or metrics['first_time_match_rate'] < 50
            pdf.create_page_1(company_name, metrics, chart_files, logo_path=logo_path, urgent=urgent_flag)
            pdf.create_page_2()
            pdf_filename = f"reports/AP_Report_{safe_name}.pdf"
            pdf.output(pdf_filename)
            pdf_filenames.append(pdf_filename)
            print(f"Generated report for {company_name}")
        df['personalisation'] = pdf_filenames
        output_file_path = 'fin_data_with_reports.xlsx'
        df.to_excel(output_file_path, index=False)
        print(f"\nSuccessfully created new file '{output_file_path}' with report links.")
    except Exception as e:
        import traceback
        print("An error occurred:")
        traceback.print_exc()

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    FILE_TO_PROCESS = os.path.join(base_dir, 'fin_data.xlsx')
    process_data_and_generate_reports(FILE_TO_PROCESS)