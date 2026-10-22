import os
import random
from datetime import datetime
from tkinter import filedialog
import customtkinter as ctk
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

print("✅ Starting application...")

# Color Palette
COLORS = {
    "primary": "#482683", #Cadbury-Purple
    "secondary": "#B9861F",
    "accent": "#FF6B6B",  # Red
    "success": "#2ECC71",  # Green
    "warning": "#F39C12", #Orange
    "dark": "#1A1A2E",  # Dark
    "light_bg": "#F8F9FA",  # Light background
    "card_bg": "#FFFFFF" ,  # White
    "text_dark": "#2C3E50",  # Dark text
    "text_light": "#7F8C8D",  # Light text
    "border": "#E1E8ED",  # Border color
    "gradient_1": "#667EEA",  # Gradient start
    "gradient_2": "#764BA2",  # Gradient end
}
## Data base and data
properties_database = [
    {
        "id": 1,
        "name": "Downtown Store A",
        "property_type": "Store",
        "floors": 2,
        "square_meters": 500,
        "parking_spaces": 10,
        "available_for": "Lease",
        "cost": 5000,
        "rating": 4.5,
        "views": 245,
        "created_at": "2026-01-16",
        "amenities": ["WiFi", "Security", "Parking"],
        "image_url": "🏪"
    },
    {
        "id": 2,
        "name": "Industrial Warehouse B",
        "property_type": "Warehouse",
        "floors": 1,
        "square_meters": 2000,
        "parking_spaces": 50,
        "available_for": "Purchase",
        "cost": 250000,
        "rating": 4.8,
        "views": 512,
        "created_at": "2023-09-10",
        "amenities": ["Security", "Loading Dock", "24/7 Access"],
        "image_url": "🏭"
    },
    {
        "id": 3,
        "name": "Restaurant Space C",
        "property_type": "Restaurant",
        "floors": 2,
        "square_meters": 300,
        "parking_spaces": 15,
        "available_for": "Both",
        "cost": 8000,
        "rating": 4.3,
        "views": 198,
        "created_at": "2024-01-12",
        "amenities": ["Kitchen Ready", "WiFi", "Parking"],
        "image_url": "🍽️"
    },
    {
        "id": 4,
        "name": "Tech Hub Office D",
        "property_type": "Office",
        "floors": 4,
        "square_meters": 1500,
        "parking_spaces": 40,
        "available_for": "Lease",
        "cost": 12000,
        "rating": 4.7,
        "views": 623,
        "created_at": "2024-12-08",
        "amenities": ["Modern Design", "WiFi", "Meeting Rooms"],
        "image_url": "🏢"
    },
    {
        "id": 5,
        "name": "Small Shop E",
        "property_type": "Store",
        "floors": 1,
        "square_meters": 200,
        "parking_spaces": 5,
        "available_for": "Both",
        "cost": 2500,
        "rating": 4.2,
        "views": 142,
        "created_at": "2025-01-20",
        "amenities": ["WiFi", "Security"],
        "image_url": "🏬"
    },
    {
        "id": 6,
        "name": "Premium Restaurant F",
        "property_type": "Restaurant",
        "floors": 3,
        "square_meters": 600,
        "parking_spaces": 30,
        "available_for": "Both",
        "cost": 15000,
        "rating": 4.9,
        "views": 789,
        "created_at": "2025-01-05",
        "amenities": ["Premium Kitchen", "WiFi", "Valet Parking"],
        "image_url": "🍴"
    },
    {
        "id": 7,
        "name": "Mega Warehouse G",
        "property_type": "Warehouse",
        "floors": 1,
        "square_meters": 5000,
        "parking_spaces": 100,
        "available_for": "Purchase",
        "cost": 500000,
        "rating": 4.6,
        "views": 934,
        "created_at": "2025-11-02",
        "amenities": ["Loading Dock", "Security", "24/7 Access"],
        "image_url": "🏗️"
    },
    {
        "id": 8,
        "name": "Modern Office Complex H",
        "property_type": "Office",
        "floors": 3,
        "square_meters": 800,
        "parking_spaces": 25,
        "available_for": "Lease",
        "cost": 6500,
        "rating": 4.4,
        "views": 567,
        "created_at": "2026-01-18",
        "amenities": ["Modern Design", "WiFi", "Cafeteria"],
        "image_url": "💼"
    },
]
## DEFAULT ADMIN
admin_accounts = {
    "Boss": "demo-admin-password",
}
## DEFAULT USER
user_accounts = {
    "demo@example.invalid": {
        "password": "demo-password",
        "name": "Stranger",
        "created_at": "2026-01-01",
        "favorites": [],
        "searches": []
    }
}
## UTILITY FUNCTIONS
def search_by_type(property_type):
    return [p for p in properties_database
            if p['property_type'].lower() == property_type.lower()]
def search_by_budget(max_budget):
    return [p for p in properties_database if p['cost'] <= max_budget]
def search_by_size(min_size):
    return [p for p in properties_database if p['square_meters'] >= min_size]
def search_by_parking(min_parking):
    return [p for p in properties_database if p['parking_spaces'] >= min_parking]
def search_by_availability(available_for):
    return [p for p in properties_database
            if p['available_for'].lower() == available_for.lower()]
def search_by_floors(num_floors):
    return [p for p in properties_database if p['floors'] == num_floors]
def search_by_rating(min_rating):
    return [p for p in properties_database if p.get('rating', 0) >= min_rating]
def get_recommended_properties(user_budget=None):

## Rule-based ranking by rating and view count (no AI model)
    recommended = sorted(properties_database,
                         key=lambda x: (x.get('rating', 0), -x.get('views', 0)),
                         reverse=True)[:5]
    if user_budget:
        recommended = [p for p in recommended if p['cost'] <= user_budget]
    return recommended
def add_new_property(name, prop_type, floors, square_meters, parking, available_for, cost):
    next_id = max([p['id'] for p in properties_database]) + 1
    new_property = {
        "id": next_id,
        "name": name,
        "property_type": prop_type,
        "floors": floors,
        "square_meters": square_meters,
        "parking_spaces": parking,
        "available_for": available_for,
        "cost": cost,
        "rating": 0,
        "views": 0,
        "created_at": datetime.now().strftime("%Y-%m-%d"),
        "amenities": [],
        "image_url": "🏠"
    }
    properties_database.append(new_property)
    return new_property

## Generate Excel report with saved dialog
def generate_excel_report(search_results, search_criteria, parent_window=None):
## Asks user where to save the file
    default_filename = f"property_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_filename,
        filetypes=[
            ("Excel Files", "*.xlsx"),
            ("All Files", "*.*")
        ],
        title="Save Report As...💭"
    )
    if not file_path:
        return None
## Create workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Property Report"
## Styling
    header_fill = PatternFill(start_color="667EEA", end_color="667EEA", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
## Title
    sheet.merge_cells('A1:J1')
    title_cell = sheet['A1']
    title_cell.value = "COMMERCIAL PROPERTY SEARCH REPORT"
    title_cell.font = Font(bold=True, size=14, color="667EEA")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    sheet.row_dimensions[1].height = 25
## Meta info
    sheet['A2'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    sheet['A3'] = f"Search Criteria: {search_criteria}"
    sheet['A4'] = f"Total Results: {len(search_results)}"

## Header
    headers = ["ID", "Name", "Type", "Floors", "Size (sqm)", "Parking",
               "Available For", "Cost (£)", "Rating", "Views"]
    for col_num, header in enumerate(headers, 1):
        cell = sheet.cell(row=6, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')
## Data rows
    for row_num, property_item in enumerate(search_results, 7):
        sheet.cell(row=row_num, column=1, value=property_item['id'])
        sheet.cell(row=row_num, column=2, value=property_item['name'])
        sheet.cell(row=row_num, column=3, value=property_item['property_type'])
        sheet.cell(row=row_num, column=4, value=property_item['floors'])
        sheet.cell(row=row_num, column=5, value=property_item['square_meters'])
        sheet.cell(row=row_num, column=6, value=property_item['parking_spaces'])
        sheet.cell(row=row_num, column=7, value=property_item['available_for'])
        sheet.cell(row=row_num, column=8, value=property_item['cost'])
        sheet.cell(row=row_num, column=9, value=property_item.get('rating', 0))
        sheet.cell(row=row_num, column=10, value=property_item.get('views', 0))

## Formating cost column
        sheet.cell(row=row_num, column=8).number_format = '£#,##0.00'
## Adding border to all cells
        for col in range(1, 11):
            sheet.cell(row=row_num, column=col).border = thin_border
## Auto-adjust column widths
    column_widths = {'A': 6, 'B': 25, 'C': 15, 'D': 8, 'E': 12,
                     'F': 10, 'G': 15, 'H': 12, 'I': 8, 'J': 8}
    for col_letter, width in column_widths.items():
        sheet.column_dimensions[col_letter].width = width

## Save the workbook
    workbook.save(file_path)
    return file_path
## Asked AI to write some taglines
TAGLINES = {
    "name": [
        "✨ Dream big, find bigger!",
        "🎯 Your perfect space awaits...",
        "💎 Make it unforgettable",
        "🚀 Location matters, we deliver!"
    ],
    "budget": [
        "💰 Your budget, our priority",
        "🎁 Find more, spend smart",
        "💵 Let's make it work!",
        "🏦 Investment in your future"
    ],
    "size": [
        "📏 Size matters! Pick wisely",
        "🏠 Space to dream bigger",
        "📐 Just the right fit",
        "🎪 More space = More possibilities"
    ],
    "parking": [
        "🅿️ Park it like it's hot!",
        "🚗 Never hunt for parking again",
        "🅿️ Spaces galore, stress evaporated",
        "🚙 Easy access for all your rides"
    ],
    "floors": [
        "🏘️ Build your empire here",
        "⬆️ Reach new heights!",
        "🏢 More levels, more potential",
        "🎯 Perfect vertical space"
    ]
}
def get_random_tagline(field_type):
## Picks random tag-line
    if field_type in TAGLINES:
        return random.choice(TAGLINES[field_type])
    return ""

## And here's the Main-Part
## MAIN APPLICATION CLASS
class PropertyManagementApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🏢 Commercial Property Management System")
        self.geometry("1920x1080")
        self.iconbitmap(default="")
        ctk.set_appearance_mode("light")
        self.current_user = None
        self.current_user_type = None
        self.current_user_email = None
        self.favorites = []
        self.cart = []
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.show_auth_screen()
## AUTH SCREEN
    def show_auth_screen(self):
        self.clear_main_frame()
        main_frame = ctk.CTkFrame(self, fg_color=COLORS["light_bg"])
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        container = ctk.CTkFrame(main_frame, fg_color="transparent")
        container.grid(row=0, column=0)
        left_side = ctk.CTkFrame(container, fg_color=COLORS["primary"], corner_radius=25)
        left_side.grid(row=0, column=0, padx=20, pady=20)
        left_side.grid_propagate(False)
        left_side.configure(width=600, height=750)
        top_spacer = ctk.CTkLabel(left_side, text="", fg_color="transparent")
        top_spacer.pack(pady=20)
## lOGO
        logo_label = ctk.CTkLabel(
            left_side,
            text="🏢",
            font=("Helvetica", 150),
            text_color=COLORS["card_bg"]
        )
        logo_label.pack(pady=15)
        brand_title = ctk.CTkLabel(
            left_side,
            text="Property Pro²",
            font=("Helvetica", 44, "bold"),
            text_color=COLORS["card_bg"]
        )
        brand_title.pack(pady=10)
        brand_subtitle = ctk.CTkLabel(
            left_side,
            text="Find Your Perfect Space",
            font=("Helvetica", 20),
            text_color=COLORS["secondary"]
        )
        brand_subtitle.pack(pady=10)
        line_label = ctk.CTkLabel(
            left_side,
            text="●━━━━━━━━━━━━━━━━━━━━━━●",
            font=("Helvetica", 12),
            text_color=COLORS["secondary"]
        )
        line_label.pack(pady=15)
        features_frame = ctk.CTkFrame(left_side, fg_color="transparent")
        features_frame.pack(pady=15, padx=50)
        features = [
            "🔍 Advanced Property Search",
            "🎯 Smart Recommendations",
            "⭐ Based on Top-Ratings",
            "🔒 Secure & Reliable",
            "📄 Professional Reports",
        ]
        for feature in features:
            feature_label = ctk.CTkLabel(
                features_frame,
                text=feature,
                font=("Helvetica", 15),
                text_color=COLORS["card_bg"],
                anchor="w"
            )
            feature_label.pack(anchor="w", pady=6)
## Statistics on branding card
        stats_frame = ctk.CTkFrame(left_side, fg_color="transparent")
        stats_frame.pack(pady=25)
        stats = [
            ("10K+", "Properties"),
            ("1M+", "Users"),
            ("24/7", "Support")
        ]
        for stat_value, stat_label in stats:
            stat_container = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stat_container.pack(side="left", padx=20)
            value_label = ctk.CTkLabel(
                stat_container,
                text=stat_value,
                font=("Helvetica", 26, "bold"),
                text_color=COLORS["card_bg"]
            )
            value_label.pack()
            label_text = ctk.CTkLabel(
                stat_container,
                text=stat_label,
                font=("Helvetica", 11),
                text_color=COLORS["secondary"]
            )
            label_text.pack()
## Part of main screen UI
        right_side = ctk.CTkFrame(container, fg_color=COLORS["card_bg"], corner_radius=25)
        right_side.grid(row=0, column=1, padx=20, pady=20)
        right_side.grid_propagate(False)
        right_side.configure(width=600, height=750)
        tab_frame = ctk.CTkFrame(right_side, fg_color=COLORS["card_bg"])
        tab_frame.pack(fill="x", padx=30, pady=30)
        self.login_tab_var = ctk.StringVar(value="login")
## Login Button
        login_button = ctk.CTkButton(
            tab_frame,
            text="🔐 LOGIN",
            font=("Helvetica", 15, "bold"),
            text_color=COLORS["primary"],
            fg_color="transparent",
            command=lambda: (self.login_tab_var.set("login"), self.update_auth_form()),
            border_width=2,
            border_color=COLORS["primary"],
            height=45
        )
        login_button.pack(side="left", fill="x", expand=True, padx=5)
## Sign-up button
        signup_button = ctk.CTkButton(
            tab_frame,
            text="✍️ SIGN UP",
            font=("Helvetica", 15, "bold"),
            text_color=COLORS["text_light"],
            fg_color="transparent",
            command=lambda: (self.login_tab_var.set("signup"), self.update_auth_form()),
            border_width=2,
            border_color=COLORS["border"],
            height=45
        )
        signup_button.pack(side="left", fill="x", expand=True, padx=5)
        self.auth_form_frame = ctk.CTkFrame(right_side, fg_color=COLORS["card_bg"])
        self.auth_form_frame.pack(fill="both", expand=True, padx=30, pady=20)
        self.update_auth_form()
## Admin/User selection
        bottom_frame = ctk.CTkFrame(right_side, fg_color=COLORS["light_bg"], corner_radius=15)
        bottom_frame.pack(fill="x", padx=30, pady=20)
        admin_label = ctk.CTkLabel(
            bottom_frame,
            text="🏪 Huh! Are you an Administrator?",
            font=("Helvetica", 13),
            text_color=COLORS["text_light"]
        )
        admin_label.pack(anchor="w", padx=20, pady=(15, 5))
        admin_button = ctk.CTkButton(
            bottom_frame,
            text="👨‍💼 Admin Login",
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["accent"],
            text_color="white",
            height=45,
            command=self.show_admin_login
        )
        admin_button.pack(fill="x", padx=20, pady=(5, 15))
    def update_auth_form(self):
        for widget in self.auth_form_frame.winfo_children():
            widget.destroy()

        if self.login_tab_var.get() == "login":
            self.create_login_form()
        else:
            self.create_signup_form()

    def create_login_form(self):
        # Create login form
        title = ctk.CTkLabel(
            self.auth_form_frame,
            text="Hey, Welcome Back! 👏",
            font=("Helvetica", 20, "bold"),
            text_color=COLORS["text_dark"]
        )
        title.pack(pady=(0, 20))

        subtitle = ctk.CTkLabel(
            self.auth_form_frame,
            text="Login to your account",
            font=("Helvetica", 13),
            text_color=COLORS["text_light"]
        )
        subtitle.pack(pady=(0, 30))

        # Login Email
        email_label = ctk.CTkLabel(
            self.auth_form_frame,
            text="📧 Email Address",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        email_label.pack(anchor="w")

        self.login_email = ctk.CTkEntry(
            self.auth_form_frame,
            placeholder_text="Enter your email...",
            font=("Helvetica", 13),
            height=45,
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        self.login_email.pack(fill="x", pady=(5, 15))
        self.login_email.insert(0, "demo@example.invalid")

        # Password
        password_label = ctk.CTkLabel(
            self.auth_form_frame,
            text="🔒 Password",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        password_label.pack(anchor="w")

        password_row = ctk.CTkFrame(self.auth_form_frame, fg_color="transparent")
        password_row.pack(fill="x", pady=(5, 20))

        self.login_password = ctk.CTkEntry(
            password_row,
            placeholder_text="Enter your password...",
            font=("Helvetica", 13),
            height=45,
            show="•",
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        self.login_password.pack(side="left", fill="x", expand=True)
        self.login_password.insert(0, "demo-password")

        login_pw_visible = {"state": False}

        def toggle_login_pw():
            login_pw_visible["state"] = not login_pw_visible["state"]
            self.login_password.configure(show="" if login_pw_visible["state"] else "•")
            eye_btn.configure(text="👀" if login_pw_visible["state"] else "👁")

        eye_btn = ctk.CTkButton(
            password_row,
            text="👁",
            width=55,
            height=45,
            fg_color=COLORS["primary"],
            text_color="white",
            command=toggle_login_pw
        )
        eye_btn.pack(side="left", padx=(10, 0))

        remember_var = ctk.BooleanVar()
        remember_check = ctk.CTkCheckBox(
            self.auth_form_frame,
            text="Remember me",
            variable=remember_var,
            font=("Helvetica", 11),
            text_color=COLORS["text_dark"]
        )
        remember_check.pack(anchor="w", pady=(0, 20))

        login_btn = ctk.CTkButton(
            self.auth_form_frame,
            text="🚀 LOGIN NOW",
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=50,
            command=self.user_login
        )
        login_btn.pack(fill="x", pady=(0, 15))

        forgot_label = ctk.CTkLabel(
            self.auth_form_frame,
            text="Forgot Password?",
            font=("Helvetica", 11),
            text_color=COLORS["secondary"]
        )
        forgot_label.pack()

    def create_signup_form(self):
        # Signup Form
        title = ctk.CTkLabel(
            self.auth_form_frame,
            text="Join Our Community!🥳",
            font=("Helvetica", 20, "bold"),
            text_color=COLORS["text_dark"]
        )
        title.pack(pady=(0, 20))

        name_label = ctk.CTkLabel(
            self.auth_form_frame,
            text="👤 Full Name",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        name_label.pack(anchor="w")

        self.signup_name = ctk.CTkEntry(
            self.auth_form_frame,
            placeholder_text="Zain Ul-Abideen",
            font=("Helvetica", 13),
            height=45,
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        self.signup_name.pack(fill="x", pady=(5, 15))

        email_label = ctk.CTkLabel(
            self.auth_form_frame,
            text="📧 Email Address",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        email_label.pack(anchor="w")

        self.signup_email = ctk.CTkEntry(
            self.auth_form_frame,
            placeholder_text="name@gmail.com",
            font=("Helvetica", 13),
            height=45,
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        self.signup_email.pack(fill="x", pady=(5, 15))

        password_label = ctk.CTkLabel(
            self.auth_form_frame,
            text="🔒 Password",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        password_label.pack(anchor="w")

        password_row = ctk.CTkFrame(self.auth_form_frame, fg_color="transparent")
        password_row.pack(fill="x", pady=(5, 15))

        self.signup_password = ctk.CTkEntry(
            password_row,
            placeholder_text="Minimum 8 characters",
            font=("Helvetica", 13),
            height=45,
            show="•",
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        self.signup_password.pack(side="left", fill="x", expand=True)

        signup_pw_visible = {"state": False}

        def toggle_signup_pw():
            signup_pw_visible["state"] = not signup_pw_visible["state"]
            self.signup_password.configure(show="" if signup_pw_visible["state"] else "•")
            eye_btn.configure(text="🙈" if signup_pw_visible["state"] else "👁")

        eye_btn = ctk.CTkButton(
            password_row,
            text="👁",
            width=55,
            height=45,
            fg_color=COLORS["primary"],
            text_color="white",
            command=toggle_signup_pw
        )
        eye_btn.pack(side="left", padx=(10, 0))

        terms_var = ctk.BooleanVar()
        terms_check = ctk.CTkCheckBox(
            self.auth_form_frame,
            text="I agree to Terms & Conditions",
            variable=terms_var,
            font=("Helvetica", 11),
            text_color=COLORS["text_dark"]
        )
        terms_check.pack(anchor="w", pady=(0, 20))

        signup_btn = ctk.CTkButton(
            self.auth_form_frame,
            text="✍️ CREATE ACCOUNT",
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["success"],
            text_color="white",
            height=50,
            command=self.user_signup
        )
        signup_btn.pack(fill="x")

    def user_login(self):
        # User Login
        email = self.login_email.get()
        password = self.login_password.get()

        if email in user_accounts and user_accounts[email]["password"] == password:
            self.current_user = user_accounts[email]["name"]
            self.current_user_email = email
            self.current_user_type = "user"
            self.show_user_dashboard()
        else:
            self.show_error("Invalid email or password! 😒")

    def user_signup(self):
        name = self.signup_name.get()
        email = self.signup_email.get()
        password = self.signup_password.get()

        if not name or not email or not password:
            self.show_error("Don't be Lazy! Fill all fields!")
            return

        if len(password) < 8:
            self.show_error("Password must be at least 8 characters!")
            return

        if email in user_accounts:
            self.show_error("Email already registered!")
            return

        user_accounts[email] = {
            "password": password,
            "name": name,
            "created_at": datetime.now().strftime("%Y-%m-%d"),
            "favorites": [],
            "searches": []
        }

        self.show_success("✅ Account created successfully!\nYou can now login.")
        self.login_tab_var.set("login")
        self.update_auth_form()
## Admin login pop-up
    def show_admin_login(self):
        admin_window = ctk.CTkToplevel(self)
        admin_window.title("Admin Login")
        admin_window.geometry("500x400")
        admin_window.resizable(False, False)
        admin_window.attributes('-topmost', True)

        title = ctk.CTkLabel(
            admin_window,
            text="👨‍💼 ADMIN AUTHENTICATION",
            font=("Helvetica", 18, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=20)

        username_label = ctk.CTkLabel(
            admin_window,
            text="👤 Username",
            font=("Helvetica", 13),
            text_color=COLORS["text_dark"]
        )
        username_label.pack(anchor="w", padx=40, pady=(10, 0))
        username_entry = ctk.CTkEntry(
            admin_window,
            placeholder_text="Enter username",
            font=("Helvetica", 12),
            height=40,
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        username_entry.pack(fill="x", padx=40, pady=(5, 15))
        username_entry.insert(0, "Boss")

        password_label = ctk.CTkLabel(
            admin_window,
            text="🔒 Password",
            font=("Helvetica", 13),
            text_color=COLORS["text_dark"]
        )
        password_label.pack(anchor="w", padx=40, pady=(10, 0))
        password_entry = ctk.CTkEntry(
            admin_window,
            placeholder_text="Enter password",
            font=("Helvetica", 12),
            height=40,
            show="•",
            fg_color=COLORS["light_bg"],
            border_color=COLORS["primary"],
            border_width=2,
            text_color=COLORS["text_dark"]
        )
        password_entry.pack(fill="x", padx=40, pady=(5, 30))
        password_entry.insert(0, "demo-admin-password")
        def admin_authenticate():
            username = username_entry.get()
            password = password_entry.get()

            if username in admin_accounts and admin_accounts[username] == password:
                self.current_user = username
                self.current_user_type = "admin"
                admin_window.destroy()
                self.show_admin_dashboard()
            else:
                self.show_error("Invalid admin credentials! 😶‍🌫️")

## Login button
        login_btn = ctk.CTkButton(
            admin_window,
            text="🚀 LOGIN",
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["accent"],
            text_color="white",
            height=45,
            command=admin_authenticate
        )
        login_btn.pack(fill="x", padx=40, pady=(0, 10))
## Cancel button
        cancel_btn = ctk.CTkButton(
            admin_window,
            text="❌ CANCEL",
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["text_light"],
            text_color="white",
            height=45,
            command=admin_window.destroy
        )
        cancel_btn.pack(fill="x", padx=40)

  ## Admin Dashboard
    def show_admin_dashboard(self):
        self.clear_main_frame()
        main_frame = ctk.CTkFrame(self, fg_color=COLORS["light_bg"])
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        header = ctk.CTkFrame(main_frame, fg_color=COLORS["primary"], height=80)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)
        header_label = ctk.CTkLabel(
            header,
            text=f"👨‍💼 Admin Dashboard - {self.current_user}",
            font=("Helvetica", 22, "bold"),
            text_color="white"
        )
        header_label.pack(side="left", padx=30, pady=20)
        logout_btn = ctk.CTkButton(
            header,
            text="🚪 Logout",
            font=("Helvetica", 12, "bold"),
            fg_color=COLORS["accent"],
            text_color="white",
            width=120,
            height=40,
            command=self.show_auth_screen
        )
        logout_btn.pack(side="right", padx=30, pady=20)
 ## Side bar
        sidebar = ctk.CTkFrame(main_frame, fg_color=COLORS["card_bg"], width=280)
        sidebar.grid(row=1, column=0, sticky="ns", padx=15, pady=15)
        sidebar.grid_propagate(False)
        sidebar_title = ctk.CTkLabel(
            sidebar,
            text="📋 MENU",
            font=("Helvetica", 16, "bold"),
            text_color=COLORS["primary"]
        )
        sidebar_title.pack(pady=20)
        menu_items = [
            ("📊 Dashboard", self.admin_dashboard_stats),
            ("📈 Analytics", self.admin_analytics),
            ("📝 All Properties", self.admin_view_all),
            ("🔍 Search Properties", self.admin_search_menu),
            ("➕ Add Property", self.admin_add_property),
            ("👥 Manage Users", self.admin_manage_users),
            ("👨‍💼 Manage Admins", self.admin_manage_accounts),
            ("📊 Export Reports", self.admin_export_reports),
            ("⚙️ Settings", self.admin_settings),
        ]
        for item_text, item_command in menu_items:
            btn = ctk.CTkButton(
                sidebar,
                text=item_text,
                command=item_command,
                fg_color="transparent",
                text_color=COLORS["text_dark"],
                font=("Helvetica", 13),
                anchor="w",
                height=50,
                hover_color=COLORS["light_bg"],
                border_color=COLORS["primary"],
                border_width=2
            )
            btn.pack(fill="x", padx=10, pady=8)
## main content
        self.content_frame = ctk.CTkFrame(main_frame, fg_color=COLORS["card_bg"], corner_radius=15)
        self.content_frame.grid(row=1, column=1, sticky="nsew", padx=15, pady=15)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.admin_dashboard_stats()

    def admin_dashboard_stats(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="📊 Dashboard Overview",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
    ## Cards
        stats = [
            ("🏠 Total Properties", len(properties_database), COLORS["primary"]),
            ("👥 Total Users", len(user_accounts), COLORS["secondary"]),
            ("👨‍💼 Admin Accounts", len(admin_accounts), COLORS["accent"]),
            ("📊 Total Views", sum(p.get('views', 0) for p in properties_database), COLORS["success"]),
        ]
        stats_frame = ctk.CTkFrame(frame, fg_color="transparent")
        stats_frame.pack(fill="x", pady=20)
        for i, (label, value, color) in enumerate(stats):
            stat_card = ctk.CTkFrame(stats_frame, fg_color=COLORS["light_bg"], corner_radius=10)
            stat_card.grid(row=0, column=i, sticky="nsew", padx=10)
            stat_label = ctk.CTkLabel(
                stat_card,
                text=label,
                font=("Helvetica", 14, "bold"),
                text_color=COLORS["text_dark"]
            )
            stat_label.pack(pady=(15, 5), padx=15)
            stat_value = ctk.CTkLabel(
                stat_card,
                text=str(value),
                font=("Helvetica", 28, "bold"),
                text_color=color
            )
            stat_value.pack(pady=(0, 15))
        featured_label = ctk.CTkLabel(
            frame,
            text="⭐ Top Rated Properties",
            font=("Helvetica", 18, "bold"),
            text_color=COLORS["primary"]
        )
        featured_label.pack(pady=(30, 20))
        top_properties = sorted(properties_database, key=lambda x: x.get('rating', 0), reverse=True)[:3]
        for prop in top_properties:
            self.create_admin_property_card(frame, prop)
## Add property card for admin
    def create_admin_property_card(self, parent, prop):
        card = ctk.CTkFrame(parent, fg_color=COLORS["light_bg"], corner_radius=10)
        card.pack(fill="x", padx=20, pady=10)
        top_frame = ctk.CTkFrame(card, fg_color="transparent")
        top_frame.pack(fill="x", padx=20, pady=(15, 10))
        name_label = ctk.CTkLabel(
            top_frame,
            text=f"{prop['image_url']} {prop['name']}",
            font=("Helvetica", 15, "bold"),
            text_color=COLORS["text_dark"]
        )
        name_label.pack(anchor="w", side="left")
        rating_label = ctk.CTkLabel(
            top_frame,
            text=f"⭐ {prop.get('rating', 0)}/5.0",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["warning"]
        )
        rating_label.pack(anchor="e", side="right")
        details_frame = ctk.CTkFrame(card, fg_color="transparent")
        details_frame.pack(fill="x", padx=20, pady=10)
        details_text = f"🏢 {prop['property_type']} | 📐 {prop['square_meters']}sqm | 🚗 {prop['parking_spaces']} spaces | 💰 £{prop['cost']:,.0f}"
        details_label = ctk.CTkLabel(
            details_frame,
            text=details_text,
            font=("Helvetica", 12),
            text_color=COLORS["text_light"]
        )
        details_label.pack(anchor="w")

    def admin_analytics(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="📈 Analytics & Insights",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
## Property type breakdown
        type_breakdown = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        type_breakdown.pack(fill="x", pady=20)
        type_label = ctk.CTkLabel(
            type_breakdown,
            text="Property Type Distribution",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["text_dark"]
        )
        type_label.pack(pady=(15, 10), padx=20)
        types = {}
        for prop in properties_database:
            prop_type = prop['property_type']
            types[prop_type] = types.get(prop_type, 0) + 1
        for prop_type, count in types.items():
            percentage = (count / len(properties_database)) * 100
            bar_frame = ctk.CTkFrame(type_breakdown, fg_color="transparent")
            bar_frame.pack(fill="x", padx=20, pady=8)
            type_name = ctk.CTkLabel(
                bar_frame,
                text=f"{prop_type}",
                font=("Helvetica", 12),
                text_color=COLORS["text_dark"],
                width=100
            )
            type_name.pack(side="left")
            bar = ctk.CTkProgressBar(
                bar_frame,
                fg_color=COLORS["border"],
                progress_color=COLORS["primary"]
            )
            bar.pack(side="left", fill="x", expand=True, padx=10)

 # Anyltic Progress bar (bug-fixed)
            bar.set(percentage / 100)
            bar.pack(side="left", fill="x", expand=True, padx=10)
            count_label = ctk.CTkLabel(
                bar_frame,
                text=f"{count} ({percentage:.0f}%)",
                font=("Helvetica", 11),
                text_color=COLORS["text_light"],
                width=100
            )
            count_label.pack(side="left", padx=10)

        metrics_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        metrics_frame.pack(fill="x", pady=20)
        metrics_label = ctk.CTkLabel(
            metrics_frame,
            text="Key Metrics",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["text_dark"]
        )
        metrics_label.pack(pady=(15, 10), padx=20)
        avg_rating = sum(p.get('rating', 0) for p in properties_database) / len(properties_database)
        avg_size = sum(p['square_meters'] for p in properties_database) / len(properties_database)
        avg_cost = sum(p['cost'] for p in properties_database) / len(properties_database)
        metrics = [
            ("Average Rating", f"{avg_rating:.2f}/5.0", COLORS["warning"]),
            ("Average Size", f"{avg_size:.0f} sqm", COLORS["primary"]),
            ("Average Cost", f"£{avg_cost:,.0f}", COLORS["success"]),
        ]
        metrics_inner = ctk.CTkFrame(metrics_frame, fg_color="transparent")
        metrics_inner.pack(fill="x", padx=20, pady=(0, 15))
        for metric_name, metric_value, color in metrics:
            metric_box = ctk.CTkFrame(metrics_inner, fg_color=COLORS["card_bg"], border_color=color, border_width=2,
                                      corner_radius=8)
            metric_box.pack(side="left", fill="x", expand=True, padx=10)
            m_name = ctk.CTkLabel(
                metric_box,
                text=metric_name,
                font=("Helvetica", 11),
                text_color=COLORS["text_light"]
            )
            m_name.pack(pady=(10, 0))
            m_value = ctk.CTkLabel(
                metric_box,
                text=metric_value,
                font=("Helvetica", 16, "bold"),
                text_color=color
            )
            m_value.pack(pady=(5, 10))
## View all properties you add as an admin
    def admin_view_all(self):
        self.display_admin_properties(properties_database, "All Properties")

    def admin_search_menu(self):
        self.clear_content_frame()

        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="🔍 Search Properties",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
        search_options = [
            ("Search by Type", self.admin_search_type),
            ("Search by Budget", self.admin_search_budget),
            ("Search by Size", self.admin_search_size),
            ("Search by Parking", self.admin_search_parking),
            ("Search by Rating", self.admin_search_rating),
            ("Search by Availability", self.admin_search_availability),
        ]
        for option_text, option_command in search_options:
            btn = ctk.CTkButton(
                frame,
                text=option_text,
                font=("Helvetica", 14, "bold"),
                fg_color=COLORS["primary"],
                text_color="white",
                height=50,
                command=option_command
            )
            btn.pack(fill="x", padx=20, pady=10)
    def admin_search_type(self):
        self.show_admin_search_dialog(
            "Search by Property Type",
            "Enter property type (Store/Warehouse/Restaurant/Office/Other):",
            lambda x: search_by_type(x),
            "Type"
        )
    def admin_search_budget(self):
        self.show_admin_search_dialog_numeric(
            "Search by Budget",
            "Enter maximum budget (£):",
            lambda x: search_by_budget(x),
            "Budget"
        )
    def admin_search_size(self):
        self.show_admin_search_dialog_numeric(
            "Search by Size",
            "Enter minimum square meters:",
            lambda x: search_by_size(x),
            "Size"
        )
    def admin_search_parking(self):
        self.show_admin_search_dialog_numeric(
            "Search by Parking",
            "Enter minimum parking spaces:",
            lambda x: search_by_parking(int(x)),
            "Parking"
        )
    def admin_search_rating(self):
        self.show_admin_search_dialog_numeric(
            "Search by Rating",
            "Enter minimum rating (0-5):",
            lambda x: search_by_rating(float(x)),
            "Rating"
        )
    def admin_search_availability(self):
        self.show_admin_search_dialog(
            "Search by Availability",
            "Enter availability (Lease/Purchase/Both):",
            lambda x: search_by_availability(x),
            "Availability"
        )
## add new property form
    def admin_add_property(self):
        self.clear_content_frame()
        form_frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        form_frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            form_frame,
            text="➕ Add New Property",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
        fields = {}
        field_configs = [
            ("Property Name", "name", "text", "Dream property name awaits..."),
            ("Property Type", "property_type", "dropdown", ["Store", "Warehouse", "Restaurant", "Office", "Other"]),
            ("Number of Floors", "floors", "dropdown", ["1", "2", "3", "4"]),
            ("Square Meters", "square_meters", "number", "📏 Size matters!"),
            ("Parking Spaces", "parking_spaces", "number", "🚗 Park it like it's hot!"),
            ("Available For", "available_for", "dropdown", ["Lease", "Purchase", "Both"]),
            ("Cost (£)", "cost", "number", "💰 Investment in your future"),
        ]
        for config in field_configs:
            if config[2] == "dropdown":
                label = ctk.CTkLabel(
                    form_frame,
                    text=config[0] + ":",
                    font=("Helvetica", 13, "bold"),
                    text_color=COLORS["text_dark"]
                )
                label.pack(anchor="w", padx=20, pady=(15, 5))
                combo = ctk.CTkComboBox(
                    form_frame,
                    values=config[3],
                    state="readonly",
                    fg_color=COLORS["light_bg"],
                    button_color=COLORS["primary"],
                    font=("Helvetica", 12),
                    text_color=COLORS["text_dark"],
                    border_color=COLORS["primary"],
                    border_width=2,
                    height=40
                )
                combo.pack(fill="x", padx=20, pady=(0, 5))
                fields[config[1]] = combo
            else:
                label = ctk.CTkLabel(
                    form_frame,
                    text=f"{config[0]}: {config[3]}",
                    font=("Helvetica", 13, "bold"),
                    text_color=COLORS["text_dark"]
                )
                label.pack(anchor="w", padx=20, pady=(15, 5))
                entry = ctk.CTkEntry(
                    form_frame,
                    placeholder_text=config[3],
                    fg_color=COLORS["light_bg"],
                    font=("Helvetica", 12),
                    text_color=COLORS["text_dark"],
                    border_color=COLORS["primary"],
                    border_width=2,
                    height=40
                )
                entry.pack(fill="x", padx=20, pady=(0, 5))
                fields[config[1]] = entry
        def submit_property():
            try:
                new_prop = add_new_property(
                    name=fields['name'].get(),
                    prop_type=fields['property_type'].get(),
                    floors=int(fields['floors'].get()),
                    square_meters=float(fields['square_meters'].get()),
                    parking=int(fields['parking_spaces'].get()),
                    available_for=fields['available_for'].get(),
                    cost=float(fields['cost'].get())
                )
                self.show_success(f"✅ Property '{new_prop['name']}' added successfully!")
                self.admin_view_all()
            except Exception as e:
                self.show_error(f"Error: {str(e)}")
        submit_btn = ctk.CTkButton(
            form_frame,
            text="✅ ADD PROPERTY",
            command=submit_property,
            font=("Helvetica", 14, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=50
        )
        submit_btn.pack(fill="x", padx=20, pady=30)
## manage user acc
    def admin_manage_users(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="👥 Manage User Accounts",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
## user list
        for email, user_info in user_accounts.items():
            user_card = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
            user_card.pack(fill="x", padx=20, pady=10)
            user_header = ctk.CTkFrame(user_card, fg_color="transparent")
            user_header.pack(fill="x", padx=15, pady=(15, 10))
            name_label = ctk.CTkLabel(
                user_header,
                text=f"👤 {user_info['name']}",
                font=("Helvetica", 14, "bold"),
                text_color=COLORS["text_dark"]
            )
            name_label.pack(anchor="w", side="left")
            email_label = ctk.CTkLabel(
                user_header,
                text=email,
                font=("Helvetica", 12),
                text_color=COLORS["text_light"]
            )
            email_label.pack(anchor="w")
            info_frame = ctk.CTkFrame(user_card, fg_color="transparent")
            info_frame.pack(fill="x", padx=15, pady=(0, 15))
            info_text = f"Joined: {user_info['created_at']} | Favorites: {len(user_info.get('favorites', []))} | Searches: {len(user_info.get('searches', []))}"
            info_label = ctk.CTkLabel(
                info_frame,
                text=info_text,
                font=("Helvetica", 11),
                text_color=COLORS["text_light"]
            )
            info_label.pack(anchor="w")

    def admin_manage_accounts(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="👨‍💼 Manage Admin Accounts",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
## Current admins
        admins_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        admins_frame.pack(fill="x", padx=20, pady=10)
        admins_label = ctk.CTkLabel(
            admins_frame,
            text="Current Admin Accounts",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["primary"]
        )
        admins_label.pack(anchor="w", padx=20, pady=(15, 10))
        for username in admin_accounts.keys():
            admin_item = ctk.CTkLabel(
                admins_frame,
                text=f"  👤 {username}",
                font=("Helvetica", 12),
                text_color=COLORS["text_dark"]
            )
            admin_item.pack(anchor="w", padx=40, pady=5)
## Create new admin
        add_admin_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        add_admin_frame.pack(fill="x", padx=20, pady=20)
        add_label = ctk.CTkLabel(
            add_admin_frame,
            text="➕ Add New Admin Account",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["primary"]
        )
        add_label.pack(anchor="w", padx=20, pady=(15, 15))
        username_label = ctk.CTkLabel(
            add_admin_frame,
            text="Username:",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        username_label.pack(anchor="w", padx=20, pady=(5, 0))
        username_entry = ctk.CTkEntry(
            add_admin_frame,
            placeholder_text="Enter username",
            fg_color=COLORS["card_bg"],
            font=("Helvetica", 12),
            text_color=COLORS["text_dark"],
            border_color=COLORS["primary"],
            border_width=2,
            height=40
        )
        username_entry.pack(fill="x", padx=20, pady=(0, 15))
        password_label = ctk.CTkLabel(
            add_admin_frame,
            text="Password:",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["text_dark"]
        )
        password_label.pack(anchor="w", padx=20, pady=(5, 0))
        password_entry = ctk.CTkEntry(
            add_admin_frame,
            placeholder_text="Enter password",
            show="•",
            fg_color=COLORS["card_bg"],
            font=("Helvetica", 12),
            text_color=COLORS["text_dark"],
            border_color=COLORS["primary"],
            border_width=2,
            height=40
        )
        password_entry.pack(fill="x", padx=20, pady=(0, 20))
        def add_admin():
            username = username_entry.get()
            password = password_entry.get()
            if not username or not password:
                self.show_error("Please fill all fields! 😖")
                return
            if username in admin_accounts:
                self.show_error("Username already exists! 😖")
                return
            admin_accounts[username] = password
            self.show_success(f"✅ Admin account '{username}' created!")
            username_entry.delete(0, "end")
            password_entry.delete(0, "end")
            self.admin_manage_accounts()

        add_btn = ctk.CTkButton(
            add_admin_frame,
            text="➕ Add Admin",
            command=add_admin,
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["success"],
            text_color="white",
            height=45
        )
        add_btn.pack(fill="x", padx=20, pady=(0, 20))
## Export reports as an admin
    def admin_export_reports(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="📊 Export Reports",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
        export_options = [
            ("📋 All Properties Report", lambda: self.export_report(properties_database, "All Properties")),
            ("⭐ Top Rated Properties",
             lambda: self.export_report(sorted(properties_database, key=lambda x: x.get('rating', 0), reverse=True)[:5],
                                        "Top Rated")),
            ("💰 Properties by Price",
             lambda: self.export_report(sorted(properties_database, key=lambda x: x['cost'], reverse=True),
                                        "By Price")),
            ("📏 Properties by Size",
             lambda: self.export_report(sorted(properties_database, key=lambda x: x['square_meters'], reverse=True),
                                        "By Size")),
        ]
        for option_text, option_command in export_options:
            btn = ctk.CTkButton(
                frame,
                text=option_text,
                font=("Helvetica", 14, "bold"),
                fg_color=COLORS["primary"],
                text_color="white",
                height=50,
                command=option_command
            )
            btn.pack(fill="x", padx=20, pady=10)
## Admin settings
    def admin_settings(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title = ctk.CTkLabel(
            frame,
            text="⚙️ System Settings",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 30))
        db_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        db_frame.pack(fill="x", padx=20, pady=10)
        db_title = ctk.CTkLabel(
            db_frame,
            text="📁 Database Management",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["primary"]
        )
        db_title.pack(anchor="w", padx=20, pady=(15, 10))
        def backup_db():
            self.show_error("Backup is not implemented in this coursework prototype.")
        backup_btn = ctk.CTkButton(
            db_frame,
            text="💾 Backup Database",
            command=backup_db,
            font=("Helvetica", 12, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=40
        )
        backup_btn.pack(fill="x", padx=20, pady=(0, 15))

## Get system info
        sys_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        sys_frame.pack(fill="x", padx=20, pady=10)
        sys_title = ctk.CTkLabel(
            sys_frame,
            text="ℹ️ System Information",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["primary"]
        )
        sys_title.pack(anchor="w", padx=20, pady=(15, 10))
        sys_info = [
            (f"App Version", "2.0.0"),
            (f"Total Properties", str(len(properties_database))),
            (f"Total Users", str(len(user_accounts))),
            (f"Admin Accounts", str(len(admin_accounts))),
        ]
        for info_label, info_value in sys_info:
            info_frame = ctk.CTkFrame(sys_frame, fg_color="transparent")
            info_frame.pack(fill="x", padx=20, pady=5)
            label = ctk.CTkLabel(
                info_frame,
                text=info_label + ":",
                font=("Helvetica", 12),
                text_color=COLORS["text_dark"],
                width=150,
                anchor="w"
            )
            label.pack(side="left")
            value = ctk.CTkLabel(
                info_frame,
                text=info_value,
                font=("Helvetica", 12, "bold"),
                text_color=COLORS["primary"]
            )
            value.pack(side="left")
        spacer = ctk.CTkLabel(sys_frame, text="")
        spacer.pack(pady=10)
## view Properties in admin view
    def display_admin_properties(self, properties, title):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title_label = ctk.CTkLabel(
            frame,
            text=f"📊 {title} ({len(properties)} properties)",
            font=("Helvetica", 22, "bold"),
            text_color=COLORS["primary"]
        )
        title_label.pack(pady=(0, 30))
        if not properties:
            no_results = ctk.CTkLabel(
                frame,
                text="❌ No properties found!",
                font=("Helvetica", 16),
                text_color=COLORS["accent"]
            )
            no_results.pack(pady=50)
            return
        for property_item in properties:
            self.create_admin_property_card(frame, property_item)

        export_btn = ctk.CTkButton(
            frame,
            text="📊 Export as Excel",
            command=lambda: self.export_report(properties, title),
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["success"],
            text_color="white",
            height=45
        )
        export_btn.pack(fill="x", padx=20, pady=20)
## \export report to exel with dialog
    def export_report(self, properties, title):
        try:
            filename = generate_excel_report(properties, title, self)
            if filename:
                self.show_success(f"✅ Report exported successfully!\n\nSaved to:\n{filename}")
            else:
                self.show_error("Export cancelled.")
        except Exception as e:
            self.show_error(f"Error exporting report: {str(e)}")

    def show_admin_search_dialog(self, title, prompt, search_func, search_type):
        self.clear_content_frame()
        frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title_label = ctk.CTkLabel(
            frame,
            text=f"🔍 {title}",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title_label.pack(pady=(0, 40))
        input_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=15)
        input_frame.pack(fill="x", padx=100, pady=20)
        prompt_label = ctk.CTkLabel(
            input_frame,
            text=prompt,
            font=("Helvetica", 14),
            text_color=COLORS["text_dark"]
        )
        prompt_label.pack(anchor="w", padx=30, pady=(20, 10))
        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Enter search criteria...",
            font=("Helvetica", 14),
            fg_color=COLORS["card_bg"],
            text_color=COLORS["text_dark"],
            border_color=COLORS["primary"],
            border_width=2,
            height=50
        )
        entry.pack(fill="x", padx=30, pady=15)
        entry.focus()
        def perform_search():
            search_value = entry.get()
            if not search_value:
                self.show_error("Please enter a search value!🙄")
                return

            try:
                results = search_func(search_value)
                self.display_admin_properties(results, f"Results for {search_type}: {search_value}")
            except Exception as e:
                self.show_error(f"Error: {str(e)}")

        search_btn = ctk.CTkButton(
            input_frame,
            text="🔍 SEARCH",
            command=perform_search,
            font=("Helvetica", 14, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=50
        )
        search_btn.pack(fill="x", padx=30, pady=(0, 30))
        entry.bind("<Return>", lambda e: perform_search())

    def show_admin_search_dialog_numeric(self, title, prompt, search_func, search_type):
        self.clear_content_frame()
        frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title_label = ctk.CTkLabel(
            frame,
            text=f"🔍 {title}",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title_label.pack(pady=(0, 40))
        input_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=15)
        input_frame.pack(fill="x", padx=100, pady=20)
        prompt_label = ctk.CTkLabel(
            input_frame,
            text=prompt,
            font=("Helvetica", 14),
            text_color=COLORS["text_dark"]
        )
        prompt_label.pack(anchor="w", padx=30, pady=(20, 10))
        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Enter number...",
            font=("Helvetica", 14),
            fg_color=COLORS["card_bg"],
            text_color=COLORS["text_dark"],
            border_color=COLORS["primary"],
            border_width=2,
            height=50
        )
        entry.pack(fill="x", padx=30, pady=15)
        entry.focus()
        def perform_search():
            search_value = entry.get()
            if not search_value:
                self.show_error("Please enter a value!")
                return

            try:
                results = search_func(float(search_value))
                self.display_admin_properties(results, f"Results for {search_type}: {search_value}")
            except ValueError:
                self.show_error("Please enter a valid number!")
            except Exception as e:
                self.show_error(f"Error: {str(e)}")

        search_btn = ctk.CTkButton(
            input_frame,
            text="🔍 SEARCH",
            command=perform_search,
            font=("Helvetica", 14, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=50
        )
        search_btn.pack(fill="x", padx=30, pady=(0, 30))
        entry.bind("<Return>", lambda e: perform_search())

## User Dashboard

    def show_user_dashboard(self):
        self.clear_main_frame()
        main_frame = ctk.CTkFrame(self, fg_color=COLORS["light_bg"])
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        header = ctk.CTkFrame(main_frame, fg_color=COLORS["primary"], height=90)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)
        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(fill="both", expand=True, padx=30, pady=15)
        welcome_label = ctk.CTkLabel(
            header_content,
            text=f"👋 Welcome, {self.current_user}!",
            font=("Helvetica", 22, "bold"),
            text_color="white"
        )
        welcome_label.pack(side="left", anchor="w")
        button_frame = ctk.CTkFrame(header_content, fg_color="transparent")
        button_frame.pack(side="right")
        favorites_btn = ctk.CTkButton(
            button_frame,
            text=f"❤️ Favorites ({len(self.favorites)})",
            font=("Helvetica", 11, "bold"),
            fg_color=COLORS["accent"],
            text_color="white",
            width=150,
            height=35,
            command=self.show_user_favorites
        )
        favorites_btn.pack(side="left", padx=5)
        profile_btn = ctk.CTkButton(
            button_frame,
            text="👤 Profile",
            font=("Helvetica", 11, "bold"),
            fg_color=COLORS["secondary"],
            text_color="white",
            width=100,
            height=35,
            command=self.show_user_profile
        )
        profile_btn.pack(side="left", padx=5)
## Logout button on top right corner
        logout_btn = ctk.CTkButton(
            button_frame,
            text="❎ Logout",
            font=("Helvetica", 11, "bold"),
            fg_color=COLORS["text_light"],
            text_color="white",
            width=100,
            height=35,
            command=self.show_auth_screen
        )
        logout_btn.pack(side="left", padx=5)

        sidebar = ctk.CTkFrame(main_frame, fg_color=COLORS["card_bg"], width=280)
        sidebar.grid(row=1, column=0, sticky="ns", padx=15, pady=15)
        sidebar.grid_propagate(False)
        sidebar_title = ctk.CTkLabel(
            sidebar,
            text="🔍 SEARCH",
            font=("Helvetica", 16, "bold"),
            text_color=COLORS["primary"]
        )
        sidebar_title.pack(pady=20)
        menu_items = [
            ("🏠 Home", self.user_home),
            ("📊 All Properties", self.user_view_all),
            ("⭐ Recommendations", self.user_recommendations),
            ("🏢 Search by Type", self.user_search_type),
            ("💰 Search by Budget", self.user_search_budget),
            ("📏 Search by Size", self.user_search_size),
            ("🚗 Search by Parking", self.user_search_parking),
            ("🏘️ Search by Floors", self.user_search_floors),
            ("📋 Search by Availability", self.user_search_availability),
            ("⭐ Search by Rating", self.user_search_rating),
        ]
        for item_text, item_command in menu_items:
            btn = ctk.CTkButton(
                sidebar,
                text=item_text,
                command=item_command,
                fg_color="transparent",
                text_color=COLORS["text_dark"],
                font=("Helvetica", 12),
                anchor="w",
                height=45,
                hover_color=COLORS["light_bg"],
                border_color=COLORS["primary"],
                border_width=2
            )
            btn.pack(fill="x", padx=10, pady=6)

#Main contnet
        self.content_frame = ctk.CTkFrame(main_frame, fg_color=COLORS["card_bg"], corner_radius=15)
        self.content_frame.grid(row=1, column=1, sticky="nsew", padx=15, pady=15)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self.user_home()
## User Homepage
    def user_home(self):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        welcome = ctk.CTkLabel(
            frame,
            text="🏠 Welcome Home!",
            font=("Helvetica", 26, "bold"),
            text_color=COLORS["primary"]
        )
        welcome.pack(pady=(0, 10))
        subtitle = ctk.CTkLabel(
            frame,
            text="Discover properties tailored just for you",
            font=("Helvetica", 14),
            text_color=COLORS["text_light"]
        )
        subtitle.pack(pady=(0, 30))
        quick_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=10)
        quick_frame.pack(fill="x", padx=20, pady=20)
        quick_title = ctk.CTkLabel(
            quick_frame,
            text="⚡ Quick Search",
            font=("Helvetica", 14, "bold"),
            text_color=COLORS["primary"]
        )
        quick_title.pack(anchor="w", padx=20, pady=(15, 10))
        quick_buttons_frame = ctk.CTkFrame(quick_frame, fg_color="transparent")
        quick_buttons_frame.pack(fill="x", padx=20, pady=(0, 20))
        quick_searches = [
            ("🏪 Stores", lambda: self.user_view_properties(search_by_type("Store"), "Stores")),
            ("🏭 Warehouses", lambda: self.user_view_properties(search_by_type("Warehouse"), "Warehouses")),
            ("🍽️ Restaurants", lambda: self.user_view_properties(search_by_type("Restaurant"), "Restaurants")),
            ("💼 Offices", lambda: self.user_view_properties(search_by_type("Office"), "Offices")),
        ]
        for btn_text, btn_cmd in quick_searches:
            btn = ctk.CTkButton(
                quick_buttons_frame,
                text=btn_text,
                command=btn_cmd,
                font=("Helvetica", 11, "bold"),
                fg_color=COLORS["primary"],
                text_color="white",
                height=40
            )
            btn.pack(side="left", fill="x", expand=True, padx=5)
        rec_title = ctk.CTkLabel(
            frame,
            text="⭐ Recommended for You",
            font=("Helvetica", 18, "bold"),
            text_color=COLORS["primary"]
        )
        rec_title.pack(pady=(30, 20), anchor="w", padx=20)
        recommended = get_recommended_properties()
        for prop in recommended:
            self.create_user_property_card(frame, prop)
## View all properties
    def user_view_all(self):
        self.user_view_properties(properties_database, "All Properties")

    def user_view_properties(self, props, title):
        self.clear_content_frame()
        frame = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)
        title_label = ctk.CTkLabel(
            frame,
            text=f"📊 {title} ({len(props)} results)",
            font=("Helvetica", 22, "bold"),
            text_color=COLORS["primary"]
        )
        title_label.pack(pady=(0, 30))

        if not props:
            no_results = ctk.CTkLabel(
                frame,
                text="❌ No properties found!",
                font=("Helvetica", 16),
                text_color=COLORS["accent"]
            )
            no_results.pack(pady=50)
            return

        for prop in props:
            self.create_user_property_card(frame, prop)
        export_btn = ctk.CTkButton(
            frame,
            text="📊 Export as Excel",
            command=lambda: self.export_report(props, title),
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["success"],
            text_color="white",
            height=45
        )
        export_btn.pack(fill="x", padx=20, pady=20)
## User Proprty
    def create_user_property_card(self, parent, prop):
        card = ctk.CTkFrame(parent, fg_color=COLORS["light_bg"], corner_radius=15)
        card.pack(fill="x", padx=20, pady=12)

        # Top section
        top_frame = ctk.CTkFrame(card, fg_color="transparent")
        top_frame.pack(fill="x", padx=20, pady=(15, 10))

        name_label = ctk.CTkLabel(
            top_frame,
            text=f"{prop['image_url']} {prop['name']}",
            font=("Helvetica", 16, "bold"),
            text_color=COLORS["text_dark"]
        )
        name_label.pack(anchor="w", side="left")

        rating_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        rating_frame.pack(anchor="e", side="right")
        rating_label = ctk.CTkLabel(
            rating_frame,
            text=f"⭐ {prop.get('rating', 0)}/5.0",
            font=("Helvetica", 12, "bold"),
            text_color=COLORS["warning"]
        )
        rating_label.pack(side="left", padx=5)
        views_label = ctk.CTkLabel(
            rating_frame,
            text=f"👁️ {prop.get('views', 0)} views",
            font=("Helvetica", 11),
            text_color=COLORS["text_light"]
        )
        views_label.pack(side="left", padx=5)
        type_label = ctk.CTkLabel(
            card,
            text=f"🏷️ {prop['property_type']} | Available: {prop['available_for']}",
            font=("Helvetica", 11),
            text_color=COLORS["text_light"]
        )
        type_label.pack(anchor="w", padx=20, pady=(0, 10))

## Grid
        details_frame = ctk.CTkFrame(card, fg_color="transparent")
        details_frame.pack(fill="x", padx=20, pady=10)

        details = [
            ("🏘️ Floors", str(prop['floors'])),
            ("📏 Size", f"{prop['square_meters']} sqm"),
            ("🚗 Parking", f"{prop['parking_spaces']} spaces"),
        ]

        for i, (label_text, value_text) in enumerate(details):
            detail_label = ctk.CTkLabel(
                details_frame,
                text=f"{label_text}: {value_text}",
                font=("Helvetica", 11),
                text_color=COLORS["text_dark"]
            )
            detail_label.grid(row=0, column=i, padx=20, pady=5, sticky="w")
        bottom_frame = ctk.CTkFrame(card, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=20, pady=(10, 15))

        price_label = ctk.CTkLabel(
            bottom_frame,
            text=f"💰 £{prop['cost']:,.2f}",
            font=("Helvetica", 16, "bold"),
            text_color=COLORS["success"]
        )
        price_label.pack(anchor="w", side="left")
        button_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        button_frame.pack(anchor="e", side="right")
        fav_emoji = "❤️" if prop['id'] in self.favorites else "💔"
        fav_btn = ctk.CTkButton(
            button_frame,
            text=f"{fav_emoji} Favorite",
            font=("Helvetica", 11, "bold"),
            fg_color=COLORS["accent"],
            text_color="white",
            width=120,
            height=35,
            command=lambda: self.toggle_favorite(prop['id'])
        )
        fav_btn.pack(side="left", padx=5)

## View Details btn
        details_btn = ctk.CTkButton(
            button_frame,
            text="📋 Details",
            font=("Helvetica", 11, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            width=120,
            height=35,
            command=lambda: self.show_property_details(prop)
        )
        details_btn.pack(side="left", padx=5)

    def toggle_favorite(self, property_id):
        if property_id in self.favorites:
            self.favorites.remove(property_id)
            self.show_success(f"❌ Removed from favorites")
        else:
            self.favorites.append(property_id)
            self.show_success(f"💘 Added to favorites!")
## Property Details
    def show_property_details(self, prop):
        detail_window = ctk.CTkToplevel(self)
        detail_window.title(f"{prop['name']}")
        detail_window.geometry("600x700")
        detail_window.resizable(False, False)
        detail_window.attributes('-topmost', True)
        scroll_frame = ctk.CTkScrollableFrame(detail_window, fg_color=COLORS["light_bg"])
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=20)


        title = ctk.CTkLabel(
            scroll_frame,
            text=f"{prop['image_url']} {prop['name']}",
            font=("Helvetica", 20, "bold"),
            text_color=COLORS["primary"]
        )
        title.pack(pady=(0, 20))

        details_list = [
            ("🏷️ Type", prop['property_type']),
            ("🏘️ Floors", str(prop['floors'])),
            ("📏 Size", f"{prop['square_meters']} sqm"),
            ("🚗 Parking", f"{prop['parking_spaces']} spaces"),
            ("📋 Available", prop['available_for']),
            ("💰 Cost", f"£{prop['cost']:,.2f}"),
            ("⭐ Rating", f"{prop.get('rating', 0)}/5.0"),
            ("👁️ Views", str(prop.get('views', 0))),
            ("📅 Added", prop['created_at']),
        ]
        for label, value in details_list:
            detail_frame = ctk.CTkFrame(scroll_frame, fg_color=COLORS["card_bg"], corner_radius=8)
            detail_frame.pack(fill="x", padx=10, pady=8)
            label_widget = ctk.CTkLabel(
                detail_frame,
                text=label,
                font=("Helvetica", 12, "bold"),
                text_color=COLORS["primary"],
                width=100,
                anchor="w"
            )
            label_widget.pack(side="left", padx=15, pady=12)
            value_widget = ctk.CTkLabel(
                detail_frame,
                text=value,
                font=("Helvetica", 12),
                text_color=COLORS["text_dark"]
            )
            value_widget.pack(side="left", padx=10, pady=12)


        if prop.get('amenities'):
            amenities_frame = ctk.CTkFrame(scroll_frame, fg_color=COLORS["card_bg"], corner_radius=8)
            amenities_frame.pack(fill="x", padx=10, pady=8)
            amenities_title = ctk.CTkLabel(
                amenities_frame,
                text="🎯 Amenities",
                font=("Helvetica", 12, "bold"),
                text_color=COLORS["primary"]
            )
            amenities_title.pack(anchor="w", padx=15, pady=(12, 5))
            amenities_text = ", ".join(prop['amenities'])
            amenities_label = ctk.CTkLabel(
                amenities_frame,
                text=amenities_text,
                font=("Helvetica", 11),
                text_color=COLORS["text_dark"]
            )
            amenities_label.pack(anchor="w", padx=15, pady=(0, 12))

        contact_btn = ctk.CTkButton(
            scroll_frame,
            text="📞 Contact Agent",
            font=("Helvetica", 13, "bold"),
            fg_color=COLORS["success"],
            text_color="white",
            height=45,
            command=lambda: self.show_error("Contacting an agent is not implemented in this prototype.")
        )
        contact_btn.pack(fill="x", padx=10, pady=20)
## Favourite Property
    def show_user_favorites(self):
        if not self.favorites:
            self.show_error("You don't have any favorites yet! 🥺")
            return

        favorite_props = [p for p in properties_database if p['id'] in self.favorites]
        self.user_view_properties(favorite_props, "❤️ My Favorites")

    def show_user_profile(self):
        profile_window = ctk.CTkToplevel(self)
        profile_window.title("👤 My Profile")
        profile_window.geometry("500x600")
        profile_window.resizable(False, False)
        profile_window.attributes('-topmost', True)

        frame = ctk.CTkScrollableFrame(profile_window, fg_color=COLORS["light_bg"])
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        header = ctk.CTkFrame(frame, fg_color=COLORS["card_bg"], corner_radius=15)
        header.pack(fill="x", pady=20)
        avatar = ctk.CTkLabel(
            header,
            text="👤",
            font=("Helvetica", 60),
            text_color=COLORS["primary"]
        )
        avatar.pack(pady=(20, 10))
        name = ctk.CTkLabel(
            header,
            text=self.current_user,
            font=("Helvetica", 20, "bold"),
            text_color=COLORS["text_dark"]
        )
        name.pack(pady=(0, 5))
        email = ctk.CTkLabel(
            header,
            text=self.current_user_email,
            font=("Helvetica", 13),
            text_color=COLORS["text_light"]
        )
        email.pack(pady=(0, 20))

## User Profile info
        info_frame = ctk.CTkFrame(frame, fg_color=COLORS["card_bg"], corner_radius=10)
        info_frame.pack(fill="x", pady=10)

        user_data = user_accounts[self.current_user_email]
        info_items = [
            (f"Member Since", user_data['created_at']),
            (f"Total Favorites", str(len(self.favorites))),
            (f"Account Status", "Active ✅"),
        ]
        for info_label, info_value in info_items:
            info_item = ctk.CTkFrame(info_frame, fg_color="transparent")
            info_item.pack(fill="x", padx=20, pady=10)
            label = ctk.CTkLabel(
                info_item,
                text=info_label,
                font=("Helvetica", 12),
                text_color=COLORS["text_light"],
                width=150,
                anchor="w"
            )
            label.pack(side="left")
            value = ctk.CTkLabel(
                info_item,
                text=info_value,
                font=("Helvetica", 12, "bold"),
                text_color=COLORS["primary"]
            )
            value.pack(side="left")

## Settings
        settings_frame = ctk.CTkFrame(frame, fg_color=COLORS["card_bg"], corner_radius=10)
        settings_frame.pack(fill="x", pady=10)
        settings_title = ctk.CTkLabel(
            settings_frame,
            text="⚙️ Settings",
            font=("Helvetica", 13, "bold"),
            text_color=COLORS["primary"]
        )
        settings_title.pack(anchor="w", padx=20, pady=(15, 10))
        settings_buttons = [
            ("🔐 Change Password", lambda: self.show_error("Password reset is not implemented in this prototype.")),
            ("📧 Update Email", lambda: self.show_error("Email updates are not implemented in this prototype.")),
            ("🔔 Notifications", lambda: self.show_error("Notification settings are not implemented in this prototype.")),
        ]
        for btn_text, btn_cmd in settings_buttons:
            btn = ctk.CTkButton(
                settings_frame,
                text=btn_text,
                command=btn_cmd,
                font=("Helvetica", 11),
                fg_color="transparent",
                text_color=COLORS["primary"],
                anchor="w",
                height=40,
                hover_color=COLORS["light_bg"],
                border_color=COLORS["primary"],
                border_width=2
            )
            btn.pack(fill="x", padx=20, pady=5)
        spacer = ctk.CTkLabel(settings_frame, text="")
        spacer.pack(pady=10)

    def user_recommendations(self):
        recommended = get_recommended_properties()
        self.user_view_properties(recommended, "⭐ Recommended for You")
    def user_search_type(self):
        self.show_user_search_dialog(
            "Search by Property Type",
            "Enter property type (Store/Warehouse/Restaurant/Office/Other):",
            lambda x: search_by_type(x),
            "Type"
        )
    def user_search_budget(self):
        self.show_user_search_dialog_numeric(
            "Search by Budget",
            "💰 What's your budget?",
            lambda x: search_by_budget(x),
            "Budget"
        )
    def user_search_size(self):
        self.show_user_search_dialog_numeric(
            "Search by Size",
            "📏 Minimum size in sqm?",
            lambda x: search_by_size(x),
            "Size"
        )
    def user_search_parking(self):
        self.show_user_search_dialog_numeric(
            "Search by Parking",
            "🚗 How many spaces do you need?",
            lambda x: search_by_parking(int(x)),
            "Parking"
        )
    def user_search_floors(self):
        self.show_user_search_dialog_numeric(
            "Search by Floors",
            "🏘️ How many floors?",
            lambda x: search_by_floors(int(x)),
            "Floors"
        )
    def user_search_availability(self):
        self.show_user_search_dialog(
            "Search by Availability",
            "Looking to Lease or Purchase?",
            lambda x: search_by_availability(x),
            "Availability"
        )
    def user_search_rating(self):
        self.show_user_search_dialog_numeric(
            "Search by Rating",
            "⭐ Minimum rating (0-5)?",
            lambda x: search_by_rating(float(x)),
            "Rating"
        )
    def show_user_search_dialog(self, title, prompt, search_func, search_type):
        self.clear_content_frame()

        frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)

        title_label = ctk.CTkLabel(
            frame,
            text=f"🔍 {title}",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title_label.pack(pady=(0, 40))

        input_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=15)
        input_frame.pack(fill="x", padx=100, pady=20)

        prompt_label = ctk.CTkLabel(
            input_frame,
            text=prompt,
            font=("Helvetica", 15),
            text_color=COLORS["text_dark"]
        )
        prompt_label.pack(anchor="w", padx=30, pady=(20, 10))

        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Enter search criteria...",
            font=("Helvetica", 14),
            fg_color=COLORS["card_bg"],
            text_color=COLORS["text_dark"],
            border_color=COLORS["primary"],
            border_width=2,
            height=50
        )
        entry.pack(fill="x", padx=30, pady=15)
        entry.focus()

        def perform_search():
            search_value = entry.get()
            if not search_value:
                self.show_error("Please enter a search value!")
                return

            try:
                results = search_func(search_value)
                self.user_view_properties(results, f"Results for {search_type}: {search_value}")
            except Exception as e:
                self.show_error(f"Error: {str(e)}")

        search_btn = ctk.CTkButton(
            input_frame,
            text="🔍 SEARCH NOW",
            command=perform_search,
            font=("Helvetica", 14, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=50
        )
        search_btn.pack(fill="x", padx=30, pady=(0, 30))

        entry.bind("<Return>", lambda e: perform_search())

    def show_user_search_dialog_numeric(self, title, prompt, search_func, search_type):
        self.clear_content_frame()

        frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=30)

        title_label = ctk.CTkLabel(
            frame,
            text=f"🔍 {title}",
            font=("Helvetica", 24, "bold"),
            text_color=COLORS["primary"]
        )
        title_label.pack(pady=(0, 40))

        input_frame = ctk.CTkFrame(frame, fg_color=COLORS["light_bg"], corner_radius=15)
        input_frame.pack(fill="x", padx=100, pady=20)

        prompt_label = ctk.CTkLabel(
            input_frame,
            text=prompt,
            font=("Helvetica", 15),
            text_color=COLORS["text_dark"]
        )
        prompt_label.pack(anchor="w", padx=30, pady=(20, 10))

        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Enter number...",
            font=("Helvetica", 14),
            fg_color=COLORS["card_bg"],
            text_color=COLORS["text_dark"],
            border_color=COLORS["primary"],
            border_width=2,
            height=50
        )
        entry.pack(fill="x", padx=30, pady=15)
        entry.focus()

        def perform_search():
            search_value = entry.get()
            if not search_value:
                self.show_error("Please enter a value! 🙄")
                return

            try:
                results = search_func(float(search_value))
                self.user_view_properties(results, f"Results for {search_type}: {search_value}")
            except ValueError:
                self.show_error("Please enter a valid number! 🙄")
            except Exception as e:
                self.show_error(f"Error: {str(e)}")

        search_btn = ctk.CTkButton(
            input_frame,
            text="🔍 SEARCH NOW",
            command=perform_search,
            font=("Helvetica", 14, "bold"),
            fg_color=COLORS["primary"],
            text_color="white",
            height=50
        )
        search_btn.pack(fill="x", padx=30, pady=(0, 30))

        entry.bind("<Return>", lambda e: perform_search())

## Utility

    def clear_content_frame(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def clear_main_frame(self):
        for widget in self.winfo_children():
            widget.destroy()

    def show_error(self, message):
        error_window = ctk.CTkToplevel(self)
        error_window.title("❌ Error")
        error_window.geometry("450x250")
        error_window.resizable(False, False)
        error_window.attributes('-topmost', True)

        icon = ctk.CTkLabel(
            error_window,
            text="❌",
            font=("Helvetica", 60),
            text_color=COLORS["accent"]
        )
        icon.pack(pady=(20, 10))

        error_label = ctk.CTkLabel(
            error_window,
            text=message,
            font=("Helvetica", 14),
            text_color=COLORS["text_dark"]
        )
        error_label.pack(pady=20, padx=20)

        ok_btn = ctk.CTkButton(
            error_window,
            text="OK",
            command=error_window.destroy,
            font=("Helvetica", 12, "bold"),
            fg_color=COLORS["accent"],
            text_color="white",
            height=40,
            width=100
        )
        ok_btn.pack(pady=20)

    def show_success(self, message):
        success_window = ctk.CTkToplevel(self)
        success_window.title("✅ Success")
        success_window.geometry("450x250")
        success_window.resizable(False, False)
        success_window.attributes('-topmost', True)

        icon = ctk.CTkLabel(
            success_window,
            text="✅",
            font=("Helvetica", 60),
            text_color=COLORS["success"]
        )
        icon.pack(pady=(20, 10))

        success_label = ctk.CTkLabel(
            success_window,
            text=message,
            font=("Helvetica", 14),
            text_color=COLORS["text_dark"]
        )
        success_label.pack(pady=20, padx=20)

        ok_btn = ctk.CTkButton(
            success_window,
            text="OK",
            command=success_window.destroy,
            font=("Helvetica", 12, "bold"),
            fg_color=COLORS["success"],
            text_color="white",
            height=40,
            width=100
        )
        ok_btn.pack(pady=20)

## Runs application
if __name__ == "__main__":
    app = PropertyManagementApp()
    app.mainloop()