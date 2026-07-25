import win32com.client as win32
import sys
import os

class create_outlook_mail:
    def create_outlook_email(to_recipients, subject, body, cc_recipients=None, attachments=None):

        try:
            # Start Outlook application
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = Mail item
            
            # Handle recipients
            if isinstance(to_recipients, list):
                mail.To = "; ".join(to_recipients)
            else:
                mail.To = to_recipients

            if cc_recipients:
                if isinstance(cc_recipients, list):
                    mail.CC = "; ".join(cc_recipients)
                else:
                    mail.CC = cc_recipients

            # Set subject and body
            mail.Subject = subject
            # signature_html = create_outlook_mail.get_outlook_signature()
            mail.HTMLBody = body # + "<br><br>" + create_outlook_mail.get_outlook_signature() # For HTML body, use mail.HTMLBody

            # Add attachments if provided
            if attachments:
                for file_path in attachments:
                    try:
                        mail.Attachments.Add(file_path)
                    except Exception as e:
                        print(f"Could not attach file {file_path}: {e}")

            # Display the email (False = non-modal, True = modal)
            mail.Display(False)

            print("Email created and opened in Outlook.")
        except Exception as e:
            print(f"Error creating Outlook email: {e}")
            sys.exit(1)

    def get_outlook_signature(signature_name=None):
        """
        Reads the default or specified Outlook signature HTML file.
        If signature_name is None, tries to detect the default signature.
        """
        sig_dir = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures')

        if not os.path.exists(sig_dir):
            raise FileNotFoundError("Outlook signature folder not found.")

        # If no signature name provided, pick the first .htm file
        if signature_name is None:
            sig_files = [f for f in os.listdir(sig_dir) if f.lower().endswith('.htm')]
            if not sig_files:
                raise FileNotFoundError("No Outlook HTML signatures found.")
            signature_name = os.path.splitext(sig_files[0])[0]

        sig_path = os.path.join(sig_dir, f"{signature_name}.htm")
        if not os.path.exists(sig_path):
            raise FileNotFoundError(f"Signature '{signature_name}' not found.")

        # Read HTML content
        with open(sig_path, 'r', encoding='utf-8', errors='ignore') as f:
            signature_html = f.read()
        
        return signature_html
    
    def create_html_body(message_text, table, end_text):
        html_table = "<table border='1' style='border-collapse:collapse;'>"
        html_table += "<tr>"
        tag = "th"
        html_table += f"<{tag} style='padding:5px;'>Pos</{tag}>"
        for columnName in list(table.columns):
            tag = "th"
            html_table += f"<{tag} style='padding:5px;'>{columnName}</{tag}>"
        for i, row in enumerate(table.itertuples(index=False)):
            html_table += "<tr>"
            tag = "th"
            html_table += f"<{tag} style='padding:5px;'>{i+1}</{tag}>"
            for cell in row:
                tag = "td" # tag = "th" # if i == 0 else "td"
                html_table += f"<{tag} style='padding:5px;'>{cell}</{tag}>"
            html_table += "</tr>"
        html_table += "</table>"
        
        html_piecetable = "<html><body>\n"
        for line in message_text.splitlines():
            html_piecetable += "<p>"
            html_piecetable += line
            html_piecetable += "</p>\n"
        html_piecetable += html_table + "\n"
        for line in end_text.splitlines():
            html_piecetable += "<p>"
            html_piecetable += line
            html_piecetable += "</p>\n"
        html_piecetable += "</body></html>\n"
        
        return html_piecetable
    
    