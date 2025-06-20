import os
import tkinter as tk
from tkinter import messagebox
from docx2pdf import convert  
from docxtpl import DocxTemplate
from tkinter import filedialog
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# --- Global List of Base Questions ---
questions = [
    "Client Name", "Project Name", "Your Name", "Date Due", "Area",
    "Rev Comment", "Office Location", "Office Address", "Office Phone", "Client Address", "Client Phone"
]

# --- Extra Fields Based on Report Type ---
report_followups = {
    "Backup Generator": ["Fuel"],
    "Batteries and Chargers": ["SEF", "SIF", "TF", "WIF", "OC", "WE", "IPC", "ESIF", "Ceiling_load"],
    "Breaker": ["Voltage", "Current", "Type"],
    "Breaker with IGS": ["Voltage", "Current", "Type"],
    "Breaker with cap": ["Voltage", "Current", "Type"],
    "High Voltage Breaker": ["Voltage", "Current"],
    "Extra High Voltage Breaker": ["Voltage", "Current"],
    "Capacitor Banks and Switchers": [""],
    "CCVT": ["Voltage"],
    "CIT": ["Voltage"],
    "Construction Spec": [""],
    "Control Cables": ["Ceiling_load"],
    "Control Enclosure": ["SEF", "SIF", "TF", "WIF", "OC", "WE", "IPC", "ESIF", "Ceiling_load"],
    "CT": ["Voltage"],
    "Disconnect Switches": [""],
    "MPT": [""],
    "Neutral Ground Reactors": [""],
    "PT": ["Voltage"],
    "SST": [""],
    "Surge Arresters": [""],
    "Testing and Commissioning": ["Type"]
}

report_types = [
    "Backup Generator", "Batteries and Chargers", "Breaker", "Breaker with IGS", "Breaker with cap",
    "High Voltage Breaker", "Extra High Voltage Breaker", "Capacitor Banks and Switchers", "CCVT", "CIT",
    "Construction Spec", "Control Cables", "Control Enclosure", "CT", "Disconnect Switches",
    "MPT", "Neutral Ground Reactors", "PT",
    "SST", "Surge Arresters", "Testing and Commissioning"
]

report_type_fields = {
    "Backup Generator": ["Fuel", "SEF", "Ceiling_load"],
    "Batteries and Chargers": ["SIF", "TF", "WIF", "OC", "WE", "IPC", "ESIF"],
    "Breaker": ["Voltage", "Current", "Type"],
    "Breaker with IGS": ["Voltage", "Current", "Type"],
    "Breaker with cap": ["Voltage", "Current", "Type"],
    "High Voltage Breaker": ["Voltage", "Current"],
    "Extra High Voltage Breaker": ["Voltage", "Current"],
    "Control Cables": ["Quantity"],
    "Control Enclosure": ["Ceiling_load"],
    "Capacitor Banks and Switchers": ["Type"],
    "SST": ["Type"],
    "Surge Arresters": ["Type"],
    "CT": ["Voltage"],
    "PT": ["Voltage"],
    "CIT": ["Voltage"],
    "CCVT": ["Voltage"]
}

field_help_texts = {
    "Client Name": "Company or customer name.",
    "Project Name": "Internal or external project title.",
    "Your Name": "Person filling out the report.",
    "Date Due": "Date the report must be completed.",
    "Area": "Location of the installation, County, State",
    "Rev Comment": "Any relevant revision notes.",
    "Office Location": "Select office to auto-fill address/phone.",
    "Type": "Choose between Solar or Wind.",
    "Voltage": "System voltage rating (e.g., 480V).",
    "Current": "Expected operating current.",
    "Ceiling_load": "Structural load at ceiling attachment point."
}

field_help_texts.update({
    "Fuel": "Type of fuel (e.g., Diesel, Gasoline)",
    "SEF": "Snow Exposure Factor (C)",
    "SIF": "Snow Importance Factor (I)",
    "TF": "Thermal Factor (C)",
    "WIF": "Wind Importance Factor",
    "OC": "Occupancy Category",
    "WE": "Wind Exposure",
    "IPC": "Internal Pressure Coefficient (+/-)",
    "ESIF": "Seismic Importance Factor",
    "Ceiling_load": "Max load on ceiling mount (PSF)",
    "Voltage": "System voltage (kV)",
    "Current": "Rated current in amps",
    "Type": "Choose between Solar or Wind",
    "Quantity": "Number of units installed",
})

office_info = {
    "Appleton": ("1 Systems Drive, Appleton, WI 54914", "920-735-6900"),
    "Austin": ("8701 N. Mopac Expy, Suite 320, Austin, TX 78759", "512-485-0831"),
    "Brainerd": ("13341 Cypress Drive, Suite 101, Baxter, MN 56425", "320-253-9495"),
    "Charlotte": ("10130 Perimeter Parkway, Suite 250, Charlotte, NC 28216", "888-937-5150"),
    "Dallas": ("7557 Rambler Road, Suite 1400, Dallas, TX 75231", "972-235-3031"),
    "Denver": ("10170 Church Ranch Way, Suite 201, Westminster, CO 80021", "720-531-8350"),
    "Englewood": ("10333 E. Dry Creek Road, Suite 400, Englewood, CO 80112", "720-482-9526"),
    "Fresno": ("7110 N Fresno Street, Suite 160, Fresno, CA 93720", "559-451-0395"),
    "Fort Worth": ("500 West 7th Street, Suite 1300, Fort Worth, TX 76102", "817-953-2777"),
    "Frisco": ("11000 Frisco St. Suite 400, Frisco, TX 75033", "469-213-1800"),
    "Houston": ("20329 State Hwy 249, Suite 350, Houston, TX 77070", "281-883-0103"),
    "Kansas City": ("12900 Foster St., Suite 120, Overland Park, KS 66213", "913-851-4492"),
    "Las Vegas": ("5725 Badura Ave., Suite 100, Las Vegas, NV 89118", "702-284-5300"),
    "Madison": ("8401 Greenway Blvd., Suite 400, Middleton, WI 53562", "608-821-6600"),
    "Merced": ("388 E Yosemite Avenue, Suite 200F, Merced, CA 95340", "209-571-1765"),
    "Minneapolis": ("12701 Whitewater Drive, Suite 300, Minnetonka, MN 55343", "952-937-5150"),
    "Modesto": ("1165 Scenic Drive, Suite A, Modesto, CA 95350", "209-571-1765"),
    "New River Valley": ("80 College Street, Suite H, Christiansburg, VA 24073", "540-381-4290"),
    "Orlando": ("1064 Greenwood Boulevard, Suite 260, Lake Mary, FL 32746", "321-294-4603"),
    "Philadelphia": ("1684 S. Broad Street, Suite 120, Lansdale, PA 19446", "215-855-7477"),
    "Phoenix": ("6909 East Greenway Parkway, Suite 250, Scottsdale, AZ 85254", "480-747-6558"),
    "Plano": ("2901 Dallas Parkway, Suite 400, Plano, TX 75093", "214-473-4640"),
    "Plano (HQ)": ("2805 Dallas Parkway, Suite 150, Plano, TX 75093", "214-473-4640"),
    "Pleasanton": ("6200 Stoneridge Mall Road, Suite 330, Pleasanton, CA 94588", "925-223-8340"),
    "Raleigh": ("801 Corporate Center Drive, Suite 310, Raleigh, NC 27607", "984-202-7500"),
    "Richmond": ("15871 City View Drive, Suite 200, Midlothian, VA 23113", "804-794-0571"),
    "Rochester": ("75 Thruway Park Drive, Suite A, West Henrietta, NY 14586", "888-937-5150"),
    "Roanoke": ("1208 Corporate Circle, Roanoke, VA 24018", "540-772-9580"),
    "San Antonio": ("211 North Loop 1604 East, Suite 205, San Antonio, TX 78232", "210-265-8300"),
    "Shenandoah Valley": ("104 Industry Way, Suite 102, Staunton, VA 24401", "540-248-3220"),
    "St. Cloud": ("1900 Medical Arts Ave, Suite 100, Sartell, MN 56377", "320-253-9495")
}

# Base folder path for all templates
base_path = r"\\westwoodps.local\GFS\WPS\Renewables Division\Services\Substation\Standards\Westwood\Specifications\Automated\Spec Sheets"
# Map report type name to the full template path
report_type_to_template_path = {
    "Backup Generator": os.path.join(base_path, "Equipment Spec - Backup Generator_20250502 python enabled.docx"),
    "Batteries and Chargers": os.path.join(base_path, "Equipment Spec - Battery Charger  python enabled.docx"),
    "Breaker": os.path.join(base_path, "Equipment Spec - Breaker_20250502.docx"),
    "Breaker with IGS": os.path.join(base_path, "Equipment Spec - Breaker with IGS_20250502.docx"),
    "Breaker with cap": os.path.join(base_path, "Equipment Spec - Breaker with Capcitor Bank_20250502.docx"),
    "High Voltage Breaker": os.path.join(base_path, "Equipment Spec - HV Breaker_20250331.docx"),
    "Extra High Voltage Breaker": os.path.join(base_path, "Equipment Spec - EHV Breaker_20250502.docx"),
    "Capacitor Banks and Switchers": os.path.join(base_path, "Equipment Spec - Cap Bank & Switcher_20250502.docx"),
    "CCVT": os.path.join(base_path, "Equipment Spec - CCVT_20250502.docx"),
    "CIT": os.path.join(base_path, "Equipment Spec - CIT_20250502.docx"),
    "Construction Spec": os.path.join(base_path, "Construction Specification_20240610.docx"),
    "Control Enclosure": os.path.join(base_path, "Equipment Spec - Control Enclosure_20250502.docx"),
    "CT": os.path.join(base_path, "Equipment Spec - CT_20250502.docx"),
    "Disconnect Switches": os.path.join(base_path, "Equipment Spec - Disconnect Switch_20250502.docx"),
    "MPT": os.path.join(base_path, "Equipment Spec - Main Power Transformer_20250502.docx"),
    "Neutral Ground Reactors": os.path.join(base_path, "Equipment Spec - NGR_20250502.docx"),
    "PT": os.path.join(base_path, "Equipment Spec - PT_20250502.docx"),
    "SST": os.path.join(base_path, "Equipment Spec - SST_20250502.docx"),
    "Surge Arresters": os.path.join(base_path, "Equipment Spec - Surge Arrester_20250502.docx"),
    "Testing and Commissioning": os.path.join(base_path, "2021-09-WW-Testing_and_Commissioning_Specification.docx")
}


def fill_word_template_with_docxtpl(template_path, output_path, data, image_path=None):
    doc = DocxTemplate(template_path)
    # Replace photo path string with actual image in the Word doc
    if image_path and os.path.exists(image_path):
        data["Photo"] = InlineImage(doc, image_path, width=Inches(0.6))  # Adjust size if needed
    else:
        data["Photo"] = ""  # Show nothing if image not provided

    doc.render(data)
    doc.save(output_path)

def save_txt_in_folder(project_path, folder_name, filename, content):
    # Full folder path inside project folder
    target_folder = os.path.join(project_path, folder_name)

    # Check if folder exists; if not, create it
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    # Full file path
    filepath = os.path.join(target_folder, filename)

    # Write content to the file
    with open(filepath, "w") as f:
        f.write(content)

    print(f"File saved: {filepath}")

def get_unique_filename(folder, base_name, extension=".txt"):
    """
    Returns a unique filename in the folder by adding a numbered prefix
    if needed.

    e.g. base_name = "Equipment Spec - Solar"
    returns "Equipment Spec - Solar.txt" if not exists,
    else "Equipment Spec - Solar 1.txt", "Equipment Spec - Solar 2.txt", etc.
    """
    filename = f"{base_name}{extension}"
    counter = 1
    while os.path.exists(os.path.join(folder, filename)):
        filename = f"{base_name} {counter}{extension}"
        counter += 1
    return filename

# --- Utility to Clear Frame ---
def clear_main_frame():
    for widget in main_frame.winfo_children():
        widget.destroy()

# --- Utility to Create Quit Button ---
def create_quit_button():
    # If a previous quit button exists, destroy it
    existing = getattr(root, "quit_button_frame", None)
    if existing:
        existing.destroy()

    # Create a new frame and attach it to root
    quit_frame = tk.Frame(root)
    quit_frame.place(relx=1.0, rely=0.0, anchor="ne", x=-10, y=10)  # top-right with margin

    quit_button = tk.Button(
        quit_frame,
        text="Quit",
        command=root.quit,
        font=("Arial", 10),
        bg="red",
        fg="white"
    )
    quit_button.pack()

    # Store the frame so we can remove it later if needed
    root.quit_button_frame = quit_frame 

# --- Project Input Screen ---
def show_project_input_screen():
    clear_main_frame()
    root.title("Enter Project Number or Browse")

    create_quit_button()

    tk.Label(main_frame, text="Enter Project #:", font=("Arial", 14)).pack(pady=(40, 10))
    project_number_entry = tk.Entry(main_frame, width=30, font=("Arial", 12))
    project_number_entry.pack(pady=10)

    def submit_project_number():
        project_number = project_number_entry.get().strip()
        if not project_number:
            messagebox.showerror("Error", "Please enter a project number.")
            return

        path_first = r'C:\Users\JMoyers\OneDrive - Westwood Active Directory\Desktop'  # To be updated to N:
        path_last = r'\130_Substation\1 - Design Stage\1.05 - Equipment Specifications'
        project_path = f"{path_first}\\{project_number}{path_last}"

        if not os.path.exists(project_path):
            messagebox.showerror("Error", f"Project folder not found:\n{project_path}")
            return

        show_reports_screen(project_number, project_path)

    def browse_for_project_folder():
        selected_folder = filedialog.askdirectory(title="Select Project Folder")
        if not selected_folder:
            return

        # Extract project number from selected path if possible
        project_number = os.path.basename(os.path.normpath(selected_folder))
        project_path = selected_folder

        show_reports_screen(project_number, project_path)

    tk.Button(main_frame, text="Submit", command=submit_project_number, width=15, font=("Arial", 12)).pack(pady=10)
    tk.Label(main_frame, text="Or", font=("Arial", 12)).pack(pady=(10, 0))
    tk.Button(main_frame, text="Browse for Project Folder", command=browse_for_project_folder, width=25, font=("Arial", 12)).pack(pady=10)

# --- Report Listing Screen ---
def show_reports_screen(project_number, project_path):
    clear_main_frame()
    root.title(f"Reports in Project {project_number}")

    # --- Ensure "Text Files" folder exists ---
    text_files_folder = os.path.join(project_path, "Text Files")
    if not os.path.exists(text_files_folder):
        os.makedirs(text_files_folder)

    # --- Load .txt files only from "Text Files" folder ---
    tk.Label(main_frame, text=f"Reports in Project {project_number}:", font=("Arial", 14)).pack(pady=(30, 10))
    report_files = [f for f in os.listdir(text_files_folder) if f.endswith(".txt")]

    if not report_files:
        tk.Label(main_frame, text="No reports found.", font=("Arial", 12)).pack()
    else:
        for f in report_files:
            tk.Label(main_frame, text=f"• {f}", anchor="w", font=("Arial", 12)).pack(fill="x", padx=20)

    # --- Buttons ---
    btn_frame = tk.Frame(main_frame)
    btn_frame.pack(pady=20)

    tk.Button(btn_frame, text="Add Report", command=lambda: show_add_report_form(project_number, project_path), width=15, font=("Arial", 12)).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Change Report", command=lambda: show_change_report_form(project_number, project_path), width=15, font=("Arial", 12)).pack(side="left", padx=10)
    tk.Button(main_frame, text="Back", command=show_project_input_screen, width=10, font=("Arial", 11)).pack(pady=30)

# --- Add Report Form ---
def show_add_report_form(project_number, project_path):
    clear_main_frame()
    root.title(f"Add Report - Project {project_number}")

    tk.Label(main_frame, text="Select Report Type:", font=("Arial", 12)).pack(pady=(10, 5))

    selected_type = tk.StringVar()
    selected_type.set(report_types[0])
    tk.OptionMenu(main_frame, selected_type, *report_types).pack(pady=5)

    entries = {}

    # Frame for static fields
    static_frame = tk.Frame(main_frame)
    static_frame.pack(pady=2, padx=20, fill="x")

    # Frame for dynamic fields (based on report type)
    dynamic_frame = tk.Frame(main_frame)
    dynamic_frame.pack(pady=2, padx=20, fill="x")

    def build_static_fields():
        for q in questions:
            frame = tk.Frame(static_frame)
            frame.pack(pady=2, fill="x")
            tk.Label(frame, text=q + ":", font=("Arial", 12), width=18, anchor="w").pack(side="left")

            if q == "Office Location":
                office_var = tk.StringVar()
                office_dropdown = tk.OptionMenu(frame, office_var, *office_info.keys())
                office_dropdown.config(font=("Arial", 12))
                office_dropdown.pack(side="left", fill="x", expand=True)
                entries[q] = office_var

                def update_office_details(*args):
                    office = office_var.get()
                    address, phone = office_info.get(office, ("", ""))

                    entries["Office Address"].config(state="normal")
                    entries["Office Address"].delete(0, tk.END)
                    entries["Office Address"].insert(0, address)
                    entries["Office Address"].config(state="readonly")

                    entries["Office Phone"].config(state="normal")
                    entries["Office Phone"].delete(0, tk.END)
                    entries["Office Phone"].insert(0, phone)
                    entries["Office Phone"].config(state="readonly")

                office_var.trace_add("write", update_office_details)

            elif q in ("Office Address", "Office Phone"):
                entry = tk.Entry(frame, font=("Arial", 12), width=40)
                entry.pack(side="left", fill="x", expand=True)

                # Add comment/description to the right
                comment = field_help_texts.get(q, "")
                if comment:
                    tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)

                entries[q] = entry
            else:
                entry = tk.Entry(frame, font=("Arial", 12), width=40)
                entry.pack(side="left", fill="x", expand=True)

                # Add comment/description to the right
                comment = field_help_texts.get(q, "")
                if comment:
                    tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)

                entries[q] = entry

    def update_dynamic_fields(*args):
        # First, remove any dynamic fields from the entries dict
        for key in list(entries):
            if key in report_followups.get(selected_type.get(), []):
                entries.pop(key)

        # Clear existing dynamic fields from the frame
        for widget in dynamic_frame.winfo_children():
            widget.destroy()

        # Rebuild fields for the selected report type
        followups = report_followups.get(selected_type.get(), [])
        for q in followups:
            frame = tk.Frame(dynamic_frame)
            frame.pack(pady=2, fill="x")
            if(q != ""):
                tk.Label(frame, text=q + ":", font=("Arial", 12), width=18, anchor="w").pack(side="left")

                if q == "Type":
                    type_var = tk.StringVar()
                    type_dropdown = tk.OptionMenu(frame, type_var, "solar", "wind")
                    type_dropdown.config(font=("Arial", 12))
                    type_dropdown.pack(side="left", fill="x", expand=True)
                    entries[q] = type_var
                else:
                    entry = tk.Entry(frame, font=("Arial", 12), width=40)
                    entry.pack(side="left", fill="x", expand=True)
                    entries[q] = entry

                # Add comment/description if available
                comment = field_help_texts.get(q, "")
                if comment:
                    tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)

    # Build static fields only once
    build_static_fields()

    # --- Photo Upload Field ---
    photo_var = tk.StringVar()

    photo_frame = tk.Frame(main_frame)
    photo_frame.pack(pady=2, fill="x")

    # Move label inside photo_frame and place it on top
    tk.Label(photo_frame, text="Client Logo (optional):", font=("Arial", 12)).pack(anchor="w")

    # Frame for entry and button side-by-side
    input_row = tk.Frame(photo_frame)
    input_row.pack(fill="x")

    photo_entry = tk.Entry(input_row, textvariable=photo_var, width=40, font=("Arial", 11))
    photo_entry.pack(side="left", fill="x", expand=True)

    def browse_photo():
        path = filedialog.askopenfilename(
            initialdir=r"\\westwoodps.local\GFS\WPS\Renewables Division\Services\Structural\Templates - CAD Drawings\CLIENT LOGOS",
            title="Select Logo",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.gif")]
        )
        if path:
            photo_var.set(path)

    browse_btn = tk.Button(input_row, text="Browse", command=browse_photo, font=("Arial", 10))
    browse_btn.pack(side="left", padx=5)

    entries["Photo"] = photo_var



    # Update dynamic fields on report type change
    selected_type.trace_add("write", update_dynamic_fields)
    update_dynamic_fields()

    def save_report():
        data = {}
        for key, widget in entries.items():
            try:
                if isinstance(widget, tk.StringVar):
                    data[key] = widget.get().strip()
                elif isinstance(widget, tk.Entry):
                    data[key] = widget.get().strip()
            except tk.TclError:
                continue

        report_type = selected_type.get().strip()
        if not report_type:
            messagebox.showerror("Error", "Report Type is missing.")
            return

        # Handle Voltage/Current in filename for select types
        voltage_sensitive = ["Breaker", "Breaker with IGS", "Breaker with cap", "High Voltage Breaker", "Extra High Voltage Breaker", "CCVT", "CIT", "CT", "PT"]
        current_sensitive = ["Breaker", "Breaker with IGS", "Breaker with cap", "High Voltage Breaker", "Extra High Voltage Breaker"]

        voltage_value = data.get("Voltage", "").replace("/", "-").replace(" ", "")
        current_value = data.get("Current", "").replace("/", "-").replace(" ", "")

        name_prefix = ""
        if report_type in voltage_sensitive and voltage_value:
            name_prefix += f" {voltage_value}kV "
        if report_type in current_sensitive and current_value:
            name_prefix += f"{current_value}A "

        base_filename = f"Equipment Spec - {name_prefix}{report_type}"

        # Ensure "Text Files" folder exists
        text_files_folder = os.path.join(project_path, "Text Files")
        os.makedirs(text_files_folder, exist_ok=True)

        # Handle duplicate filenames
        def get_unique_path(folder, base_name, ext):
            i = 1
            path = os.path.join(folder, base_name + ext)
            while os.path.exists(path):
                path = os.path.join(folder, f"{base_name} ({i}){ext}")
                i += 1
            return path

        txt_filepath = get_unique_path(text_files_folder, base_filename, ".txt")
        docx_filepath = get_unique_path(project_path, base_filename, ".docx")

        # Save .txt file
        with open(txt_filepath, "w") as f:
            f.write(f"Report Type: {report_type}\n")
            for k, v in data.items():
                f.write(f"{k}: {v}\n")

        # Key mapping for title block fields
        key_mapping = {
            "Report_type": "Report Type",
            "Client": "Client Name",
            "Project": "Project Name",
            "Your_Name": "Your Name",
            "Date": "Date Due",
            "Area": "Area",
            "Rev_Comment": "Rev Comment",
            "Client_address": "Client Address",
            "Client_phone": "Client Phone",
            "Office_location": "Office Location",
            "Office_address": "Office Address",
            "Office_phone": "Office Phone"
        }

        word_data = {}
        for placeholder, entry_key in key_mapping.items():
            widget = entries.get(entry_key)
            if isinstance(widget, tk.StringVar):
                word_data[placeholder] = widget.get().strip()
            elif isinstance(widget, tk.Entry):
                word_data[placeholder] = widget.get().strip()
            else:
                word_data[placeholder] = ""

        # Include other fields
        for key, value in data.items():
            if key not in key_mapping.values():
                word_data[key] = value

        template_path = report_type_to_template_path.get(report_type)
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Error", f"Template not found: {template_path}")
            return

        image_path = data.get("Photo", "")
        if not image_path:
            image_path = None

        fill_word_template_with_docxtpl(template_path, docx_filepath, word_data, image_path=image_path)

        try:
            messagebox.showinfo("Creating Reports...","Please wait until you see the success screen, this may take up to a minute.")
            convert(docx_filepath)
        except Exception as e:
            messagebox.showwarning("PDF Conversion Failed", f"Could not convert to PDF:\n{e}")

        messagebox.showinfo("Success", f"Report saved:\n{docx_filepath}")

    # --- Top Right Button Frame ---
    top_button_frame = tk.Frame(main_frame)
    top_button_frame.pack(anchor="ne", padx=20, pady=10)

    tk.Button(top_button_frame, text="Save Report", command=save_report, width=15, font=("Arial", 12)).pack(side="right", padx=5)
    tk.Button(top_button_frame, text="Back", command=lambda: show_reports_screen(project_number, project_path), width=10, font=("Arial", 11)).pack(side="right", padx=5)

# --- Change Report: File Selector ---
def show_change_report_form(project_number, project_path):
    clear_main_frame()
    root.title(f"Change Report - Project {project_number}")

    # Check and create "Text Files" folder
    text_files_folder = os.path.join(project_path, "Text Files")
    os.makedirs(text_files_folder, exist_ok=True)

    # Only include .txt files
    report_files = [f for f in os.listdir(text_files_folder) if f.endswith(".txt")]

    if not report_files:
        tk.Label(main_frame, text="No reports found.", font=("Arial", 12)).pack(pady=20)
        tk.Button(main_frame, text="Back", command=lambda: show_reports_screen(project_number, project_path),
                  font=("Arial", 12)).pack(pady=10)
        return

    tk.Label(main_frame, text="Select a report to change:", font=("Arial", 14)).pack(pady=(30, 10))

    selected_file = tk.StringVar(value=report_files[0])  # Default to first file
    dropdown = tk.OptionMenu(main_frame, selected_file, *report_files)
    dropdown.config(font=("Arial", 12), width=40)
    dropdown.pack(pady=10)

    def open_selected_report():
        filename = selected_file.get()
        filepath = os.path.join(text_files_folder, filename)

        data = {}
        with open(filepath, "r") as f:
            for line in f:
                if ":" in line:
                    key, val = line.strip().split(":", 1)
                    data[key.strip()] = val.strip()

        show_edit_report_form(project_number, project_path, filename, data)

    tk.Button(main_frame, text="Open Report", command=open_selected_report, font=("Arial", 12)).pack(pady=10)
    tk.Button(main_frame, text="Back", command=lambda: show_reports_screen(project_number, project_path), font=("Arial", 12)).pack(pady=20)

# --- Edit Report Form ---
def show_edit_report_form(project_number, project_path, filename, data):
    clear_main_frame()  # your existing function to clear main_frame widgets
    root.title(f"Edit Report - {filename}")
    entries = {}

    report_type = data.get("Report Type", "")

    # --- Report Type field (readonly) ---
    frame = tk.Frame(main_frame)
    frame.pack(pady=10, padx=20, fill="x")
    tk.Label(frame, text="Report Type:", font=("Arial", 12), width=18, anchor="w").pack(side="left")
    report_type_entry = tk.Entry(frame, font=("Arial", 12), width=40)
    report_type_entry.insert(0, report_type)
    report_type_entry.config(state="readonly")
    report_type_entry.pack(side="left", fill="x", expand=True)
    entries["Report Type"] = report_type_entry

    # --- Build fields to show based on report type ---
    fields_to_show = list(questions)  # base common questions
    # Add report-specific extra fields
    for extra_field in report_followups.get(report_type, []):
        if extra_field not in fields_to_show:
            fields_to_show.append(extra_field)
    # Also include any keys from data not already listed (to not lose any data)
    for key in data.keys():
        if key not in fields_to_show and key != "Report Type":
            fields_to_show.append(key)

    selected_office = tk.StringVar()

    # --- Create widgets for each field ---
    for q in fields_to_show:
        frame = tk.Frame(main_frame)
        frame.pack(pady=5, padx=20, fill="x")
        display_label = "Client Logo (optional)" if q == "Photo" else q
        tk.Label(frame, text=display_label + ":", font=("Arial", 12), width=18, anchor="w").pack(side="left")


        if q == "Office Location":
            office_dropdown = tk.OptionMenu(frame, selected_office, *office_info.keys())
            selected_office.set(data.get(q, ""))  # set initial selection
            office_dropdown.pack(side="left", fill="x", expand=True)
            entries[q] = selected_office

            # Update address and phone when office location changes
            def update_office_details(*args):
                office = selected_office.get()
                address, phone = office_info.get(office, ("", ""))
                if "Office Address" in entries:
                    entries["Office Address"].config(state="normal")
                    entries["Office Address"].delete(0, tk.END)
                    entries["Office Address"].insert(0, address)
                    entries["Office Address"].config(state="readonly")
                if "Office Phone" in entries:
                    entries["Office Phone"].config(state="normal")
                    entries["Office Phone"].delete(0, tk.END)
                    entries["Office Phone"].insert(0, phone)
                    entries["Office Phone"].config(state="readonly")

            selected_office.trace_add("write", update_office_details)

            # Add help label for Office Location
            comment = field_help_texts.get(q, "")
            if comment:
                tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)

        elif q in ("Office Address", "Office Phone"):
            entry = tk.Entry(frame, font=("Arial", 12), width=40)
            entry.insert(0, data.get(q, ""))
            entry.config(state="readonly")
            entry.pack(side="left", fill="x", expand=True)
            entries[q] = entry

            # Add help label
            comment = field_help_texts.get(q, "")
            if comment:
                tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)


        elif q == "Type":
                type_var = tk.StringVar()
                type_var.set(data.get(q, ""))
                type_dropdown = tk.OptionMenu(frame, type_var, "Solar", "Wind")
                type_dropdown.config(font=("Arial", 12))
                type_dropdown.pack(side="left", fill="x", expand=True)
                entries[q] = type_var

                comment = field_help_texts.get(q, "")
                if comment:
                    tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)

        else:
            entry = tk.Entry(frame, font=("Arial", 12), width=40)
            entry.insert(0, data.get(q, ""))
            entry.pack(side="left", fill="x", expand=True)
            entries[q] = entry

            # Add help label
            comment = field_help_texts.get(q, "")
            if comment:
                tk.Label(frame, text=comment, font=("Arial", 10), fg="gray").pack(side="left", padx=10)

    # Trigger initial update of office address and phone based on selected office location
    if "Office Location" in entries:
        update_office_details()

    # --- Save Button handler ---
    def save_changes():
        data = {}
        for key, widget in entries.items():
            try:
                if isinstance(widget, tk.StringVar):
                    data[key] = widget.get().strip()
                elif isinstance(widget, tk.Entry):
                    data[key] = widget.get().strip()
            except tk.TclError:
                continue

        report_type = data.get("Report Type", "").strip()
        if not report_type:
            messagebox.showerror("Error", "Report Type is missing.")
            return

        # Ensure "Text Files" folder exists inside the project folder
        text_files_folder = os.path.join(project_path, "Text Files")
        os.makedirs(text_files_folder, exist_ok=True)

        # Generate a unique filename to avoid overwriting existing reports
        filename = get_unique_filename(text_files_folder, f"Equipment Spec - {report_type}", extension=".txt")
        filepath = os.path.join(text_files_folder, filename)

        # Save the .txt file inside the "Text Files" folder
        with open(filepath, "w") as f:
            f.write(f"Report Type: {report_type}\n")
            for k, v in data.items():
                f.write(f"{k}: {v}\n")

        # Mapping for title block / known fields
        key_mapping = {
            "Report_type": "Report Type",
            "Client": "Client Name",
            "Project": "Project Name",
            "Your_Name": "Your Name",
            "Date": "Date Due",
            "Area": "Area",
            "Rev_Comment": "Rev Comment",
            "Client_address": "Client Address",
            "Client_phone": "Client Phone",
            "Office_location": "Office Location",
            "Office_address": "Office Address",
            "Office_phone": "Office Phone"
        }

        word_data = {}

        # Add mapped header fields
        for placeholder, entry_key in key_mapping.items():
            widget = entries.get(entry_key)
            if isinstance(widget, tk.StringVar):
                word_data[placeholder] = widget.get().strip()
            elif isinstance(widget, tk.Entry):
                word_data[placeholder] = widget.get().strip()
            else:
                word_data[placeholder] = ""

        # Add extra fields directly
        for key, value in data.items():
            if key not in key_mapping.values():
                word_data[key] = value

        template_path = report_type_to_template_path.get(report_type)
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Error", f"Template not found for report type: {report_type}")
            return

        # Save .docx to main project folder with matching name
        docx_path = os.path.join(project_path, filename.replace(".txt", ".docx"))

        # Handle image path from "Photo" field
        image_path = data.get("Photo", "")
        if not image_path:
            image_path = None

        # Pass image path to template renderer
        fill_word_template_with_docxtpl(template_path, docx_path, word_data, image_path=image_path)

        try:
            messagebox.showinfo("Generating","Creating Reports...\nPlease wait until you see the success screen")
            convert(docx_path)
        except Exception as e:
            messagebox.showwarning("PDF Conversion Failed", f"Could not convert to PDF:\n{e}")


        messagebox.showinfo("Success", f"Report saved:\n{docx_path}")

    # --- Top Right Button Frame ---
    top_button_frame = tk.Frame(main_frame)
    top_button_frame.pack(anchor="ne", padx=20, pady=10)

    tk.Button(top_button_frame, text="Save Changes", command=save_changes, width=15, font=("Arial", 12)).pack(side="right", padx=5)
    tk.Button(top_button_frame, text="Back", command=lambda: show_change_report_form(project_number, project_path), width=10, font=("Arial", 11)).pack(side="right", padx=5)


# --- Main Tkinter Window ---
root = tk.Tk()
root.title("Enter Project Number")
root.state('zoomed')
main_frame = tk.Frame(root)
main_frame.pack(expand=True)
show_project_input_screen()
root.mainloop()
