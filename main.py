import jinja2
from datetime import datetime
import pdfkit

# Step 1: Prepare the context for Jinja2
today_date = datetime.today().strftime("%d %b, %Y")
context = {'date': today_date}

# Step 2: Load and render the template
template_loader = jinja2.FileSystemLoader('./')
template_env = jinja2.Environment(loader=template_loader)

try:
    template = template_env.get_template('template.html')
    output_text = template.render(context)
except jinja2.TemplateNotFound:
    print("Template not found. Please ensure 'template.html' exists in the current directory.")
    exit(1)
except jinja2.TemplateError as e:
    print(f"An error occurred while rendering the template: {e}")
    exit(1)

# Step 3: Convert the rendered HTML to PDF
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

try:
    pdfkit.from_string(output_text, "sample.pdf", configuration=config,css="style.css",options= {'enable-local-file-access': None})
    print("PDF generated successfully.")
except OSError as e:
    print(f"An error occurred while generating the PDF: {e}")
    exit(1)