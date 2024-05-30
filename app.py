import tkinter as tk
from tkinter import ttk, filedialog
import csv
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
from email.mime.text import MIMEText

csv_data = []
import os


def get_deadline_date(end_date_str):
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    deadline_date = end_date - timedelta(days=1)
    while deadline_date.weekday() >= 5:
        deadline_date -= timedelta(days=1)
    deadline_date_str = deadline_date.strftime("%Y-%m-%d")
    return deadline_date_str


def choose_csv_file():
    predefined_folder = 'Application_cap/DATABASE'  # Update with the actual path to your predefined folder

    # Use the file dialog to choose a CSV file from the predefined folder
    file_path = filedialog.askopenfilename(initialdir=predefined_folder, title='Choose CSV File',
                                           filetypes=[('CSV Files', '*.csv')])
    if file_path:
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            global csv_data
            csv_data = list(csv.DictReader(csvfile))
        update_ca_combobox()


def update_ca_combobox():
    ca_combobox['values'] = sorted(set(row['CA'] for row in csv_data))


def show_functions(event):
    selected_ca = ca_combobox.get()
    function_combobox['values'] = sorted(set(row['Function'] for row in csv_data if row['CA'] == selected_ca))


def show_details():
    selected_ca = ca_combobox.get()
    selected_function = function_combobox.get()
    selected_data = [row for row in csv_data if row['CA'] == selected_ca and row['Function'] == selected_function]

    if selected_data:
        details_text.config(state=tk.NORMAL)
        details_text.delete(1.0, tk.END)
        details_text.insert(tk.END, f"Folder: {selected_data[0]['folder']}\n")
        details_text.insert(tk.END, f"EEAD Multigame and DCI: {selected_data[0]['EEAD Multigame and DCI']}\n")
        details_text.insert(tk.END, f"Name of EEA system: {selected_data[0]['Name of EEA system']}\n")
        details_text.insert(tk.END, f"System Mapping NEAR1.0: {selected_data[0]['System Mapping NEAR1.0']}\n")
        details_text.insert(tk.END, f"System Mapping NEAR1.1: {selected_data[0]['System Mapping NEAR1.1']}\n")
        details_text.insert(tk.END, f"System Mapping NEAR1.2: {selected_data[0]['System Mapping NEAR1.2']}\n")
        details_text.insert(tk.END, f"Requirement Specification: {selected_data[0]['Requirement Specification']}\n")
        details_text.insert(tk.END, f"Proofreaders SSD: {selected_data[0]['Proofreaders SSD']}\n")
        details_text.insert(tk.END, f"Proofreaders EECA: {selected_data[0]['Proofreaders EECA']}\n")
        details_text.insert(tk.END, f"IQ LINK: {selected_data[0]['IQ LINK']}\n")
        details_text.insert(tk.END, f"Checklist Link: {selected_data[0]['Checklist Link']}\n")
        details_text.config(state=tk.DISABLED)
        confirmation_label.config(text='Details shown successfully.')
    else:
        confirmation_label.config(text='No information found for the selected CA and function.')


def send_email():
    # recipient_email = recipient_entry.get()
    recipient_email = 'nizarzayaz2001@gmail.com'
    if recipient_email:
        msg = EmailMessage()
        msg['Subject'] = f'Function: {function_combobox.get()}, CA: {ca_combobox.get()}'

        smtp_server = 'smtp-relay.brevo.com'
        smtp_port = 587
        smtp_username = 'nizarelmouaquit@gmail.com'
        smtp_password = 'nizarelmouaquith@gmail.com'

        ca_name = ca_combobox.get()
        function_name = function_combobox.get()

        selected_data = [row for row in csv_data if row['CA'] == ca_name and row['Function'] == function_name]
        if not selected_data:
            confirmation_label.config(text='No information found for the selected CA and function.')
            return

        folder = selected_data[0]['folder']
        eead_multigame_dci = selected_data[0]['EEAD Multigame and DCI']
        eea_system_mapping_near10 = selected_data[0]['Name of EEA system']
        eea_system_mapping_near11 = selected_data[0]['System Mapping NEAR1.1']
        eea_system_mapping_near12 = selected_data[0]['System Mapping NEAR1.2']
        requirement_spec = selected_data[0]['Requirement Specification']
        proofreaders_ssd = selected_data[0]['Proofreaders SSD']
        proofreaders_eeca = selected_data[0]['Proofreaders EECA']
        iq_link = selected_data[0]['IQ LINK']
        checklist_link = selected_data[0]['Checklist Link']

        end_date_str = end_date_entry.get()
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()

        deadline_date = get_deadline_date(end_date_str)
        ee_arch_cr = ee_arch_entry_cr.get()

        ee_safety_cr = ee_safety_entry_cr.get()

        ee_diagnosis_cr = ee_diagnosis_entry_cr.get()

        tech_review_cr = tech_review_entry_cr.get()

        email_body = f'''
       <html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta name="generator" content="Aspose.Words for .NET 23.7.0">
    <title></title>

  </head>
  <body style="line-height: 108%;font-family: Calibri;font-size: 11pt;">
    <div>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;text-align: center;line-height: normal;font-size: 20pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold">NEA et DCEE et NT</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Hello,</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Please find below the links to version </span>
        <span style="font-weight:bold; color:#0c64c0">( input version here ) </span>
        <span>of the function </span>
        <span style="font-weight:bold; color:#0c64c0">"</span>
        <span style="font-weight:bold; color:#ff0000">{eea_system_mapping_near10}</span>
        <span style="font-weight:bold; color:#0c64c0">":</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 14pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-size:11pt">All the updated documents and </span>
        <span style="font-size:11pt; font-weight:bold">the IQ </span>
        <span style="font-size:11pt">are available in the following folder of the function</span>
        <span style="font-size:11pt; font-weight:bold; color:#0c64c0">« {eea_system_mapping_near10} » </span>
        <span style="font-size:11pt; font-weight:bold">:</span>
        <span style="font-weight:bold; -aw-import:spaces">&#xa0; </span>
        <span style="font-size:13pt; font-weight:bold; text-decoration:underline; color:#0000ff">{folder} </span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Link of EEAD Multigame and DCI :  {eead_multigame_dci}</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Name of EEA system:</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-style:italic">{eea_system_mapping_near10}</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Name of System Mapping NEAR1.0 : </span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>{eea_system_mapping_near10}</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Name of System Mapping NEAR1.1 :</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-style:italic">{eea_system_mapping_near11}</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Name of System Mapping NEAR1.2 :</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-style:italic">{eea_system_mapping_near12}</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Name of the Requirement Specification:</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>{requirement_spec}</span>
      </p>
      <p style="margin-left: 27pt;margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold; -aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="text-decoration:underline">Summary of the changes made justifying the evolution of documents :</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 11.5pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 10pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 10pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 11.5pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold; text-decoration:underline">Impacted components : </span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 11.5pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold; text-decoration:underline">Proofreaders :</span>
        <br>
        <br>
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <table cellspacing="0" cellpadding="0" style="width: 451.5pt;margin-left: 2.4pt;margin-bottom: 0pt;border: 0.75pt solid #000000;-aw-border: 0.5pt single;border-collapse: collapse;margin-top: 0pt;">
        <tr>
          <td style="width:143pt; border-bottom-style:solid; border-bottom-width:0.75pt; padding:2.75pt 2.75pt 2.75pt 2.38pt; vertical-align:top; -aw-border-bottom:0.5pt single">
            <p class="TableContents" style="margin: 0pt 0pt 8pt;margin-top: 0pt;margin-bottom: 8pt;text-align: left;line-height: 108%;widows: 0;orphans: 0;font-family: Calibri;font-size: 11pt;color: #000000;">
              <span style="-aw-import:ignore">&#xa0;</span>
            </p>
          </td>
          <td style="width:145.25pt; border-left-style:solid; border-left-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; padding:2.75pt 2.75pt 2.75pt 2.38pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single">
            <p class="TableContents" style="margin: 0pt 0pt 8pt;margin-top: 0pt;margin-bottom: 8pt;text-align: left;line-height: 108%;widows: 0;orphans: 0;font-family: Calibri;font-size: 11pt;color: #000000;">
              <span style="-aw-import:spaces">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>
              <span style="font-weight:bold; font-style:italic">EECA</span>
            </p>
          </td>
          <td style="width:146pt; border-left-style:solid; border-left-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; padding:2.75pt 2.38pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single">
            <p class="TableContents" style="margin: 0pt 0pt 8pt;margin-top: 0pt;margin-bottom: 8pt;text-align: left;line-height: 108%;widows: 0;orphans: 0;font-family: Calibri;font-size: 11pt;color: #000000;">
              <span style="-aw-import:spaces">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>
              <span style="font-weight:bold; font-style:italic; -aw-import:spaces">&#xa0;</span>
              <span style="font-weight:bold; font-style:italic">SSD</span>
            </p>
          </td>
        </tr>
        <tr>
          <td style="width:143pt; padding:2.75pt 2.75pt 2.75pt 2.38pt; vertical-align:top">
            <p class="TableContents" style="margin: 0pt 0pt 8pt;margin-top: 0pt;margin-bottom: 8pt;text-align: left;line-height: 108%;widows: 0;orphans: 0;font-family: Calibri;font-size: 11pt;color: #000000;">
              <span style="-aw-import:spaces">&#xa0; </span>
              <span style="font-weight:bold; font-style:italic; text-decoration:underline">PROOFREADERS </span>
            </p>
          </td>
          <td style="width:145.25pt; border-left-style:solid; border-left-width:0.75pt; padding:2.75pt 2.75pt 2.75pt 2.38pt; vertical-align:top; -aw-border-left:0.5pt single">
            <p class="TableContents" style="margin: 0pt 0pt 8pt;margin-top: 0pt;margin-bottom: 8pt;text-align: left;line-height: 108%;widows: 0;orphans: 0;font-family: Calibri;font-size: 11pt;color: #000000;">
              <span style="-aw-import:spaces">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>
              <span>{proofreaders_eeca}</span>
            </p>
          </td>
          <td style="width:146pt; border-left-style:solid; border-left-width:0.75pt; padding:2.75pt 2.38pt; vertical-align:top; -aw-border-left:0.5pt single">
            <p class="TableContents" style="margin: 0pt 0pt 8pt;margin-top: 0pt;margin-bottom: 8pt;text-align: left;line-height: 108%;widows: 0;orphans: 0;font-family: Calibri;font-size: 11pt;color: #000000;">
              <span style="-aw-import:spaces">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>
              <span>{proofreaders_ssd}</span>
            </p>
          </td>
        </tr>
      </table>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold; -aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-left: 27pt;margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>The different proofreaders identified must check the changes made by the designer and update the status of the corresponding questions (for example: processed) in the IQ before the review.</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold">If you have neither remark nor comment</span>
        <span>, please indicate it by adding a comment</span>
        <span style="font-weight:bold">« reviewed without remark»</span>
        <span>in the IQ file.</span>
      </p>
      <p style="margin-left: 27pt;margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold; text-decoration:underline">End date:</span>
        <span style="font-weight:bold">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>As the Technical Review is scheduled for the </span>
        <span style="font-weight:bold; color:#c82613"> {end_date} </span>
        <span>, please add your remarks</span>
        <span style="font-weight:bold"></span>
        <span style="font-weight:bold; color:#c82613">before {deadline_date} </span>
        <span>in the IQ link: </span>
        <span style="font-weight:bold; color:#2a6099">{iq_link}</span>
        <br>
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-weight:bold; color:#0000ff; -aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>The technical verifications to be carried out are: </span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <br>
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <table cellspacing="0" cellpadding="0" style="width: 455.2pt;margin-bottom: 0pt;border: 1pt solid #000000;border-collapse: collapse;margin-top: 0pt;">
        <tr style="height:23.9pt">
          <td style="width:151.75pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">Item</span>
            </p>
          </td>
          <td style="width:79pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">Check requested</span>
            </p>
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">(Yes, No, N/A)</span>
            </p>
          </td>
          <td style="width:199.45pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">Who</span>
            </p>
          </td>
        </tr>
        <tr style="height:15.1pt">
          <td style="width:151.75pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; text-decoration:underline; color:#323130">Check EE ARCHITECTURE INTEGRATION</span>
            </p>
          </td>
          <td style="width:79pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">{ee_arch_cr}</span>
            </p>
          </td>
          <td style="width:199.45pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130; -aw-import:ignore">&#xa0;</span>
            </p>
          </td>
        </tr>
        <tr style="height:15.1pt">
          <td style="width:151.75pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; text-decoration:underline; color:#323130">Check EE SAFETY &amp; CYBER</span>
            </p>
          </td>
          <td style="width:79pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">{ee_safety_cr}</span>
            </p>
          </td>
          <td style="width:199.45pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 11pt;margin: 0pt 0pt 8pt;">
              <span style="-aw-import:ignore">&#xa0;</span>
            </p>
          </td>
        </tr>
        <tr style="height:15.25pt">
          <td style="width:151.75pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; text-decoration:underline; color:#323130">Check EE DIAGNOSIS</span>
            </p>
          </td>
          <td style="width:79pt; border-right-style:solid; border-right-width:1pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">{ee_diagnosis_cr}</span>
            </p>
          </td>
          <td style="width:199.45pt; border-bottom-style:solid; border-bottom-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 11pt;margin: 0pt 0pt 8pt;">
              <span style="-aw-import:ignore">&#xa0;</span>
            </p>
          </td>
        </tr>
        <tr style="height:15.25pt">
          <td style="width:151.75pt; border-right-style:solid; border-right-width:1pt; padding:4pt 3.5pt; vertical-align:top; background-color:#deeaf6">
            <p style="margin-bottom: 0pt;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; text-decoration:underline; color:#323130">Check Technical Review Quality</span>
            </p>
          </td>
          <td style="width:79pt; border-right-style:solid; border-right-width:1pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 10pt;margin: 0pt 0pt 8pt;">
              <span style="font-weight:bold; color:#323130">{tech_review_cr}</span>
            </p>
          </td>
          <td style="width:199.45pt; padding:4pt 3.5pt 4pt 4pt; vertical-align:top; background-color:#ffffff">
            <p style="margin-bottom: 0pt;text-align: center;widows: 0;orphans: 0;font-size: 11pt;margin: 0pt 0pt 8pt;">
              <span style="-aw-import:ignore">&#xa0;</span>
            </p>
          </td>
        </tr>
      </table>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 11.5pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="color:#242424; -aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Checklist Link: {checklist_link}</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 9pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-style:italic">The EEAD / NT are transmitted to the proofreaders defined in the invitation file for the technical reviews:</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 9pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-style:italic">If you are not concerned, please unsubscribe.</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-size:9pt; font-style:italic">If you have not received this email directly and that you are concerned, please subscribe.</span>
        <span style="font-size:9pt">&#xa0;</span>
        <a href="http://docinfogroupe.inetpsa.com/ead/doc/ref.01991_16_00069/v.vc/fiche" target="_blank" style="text-decoration:none">
          <span style="font-size:9pt; font-weight:bold; text-decoration:underline; color:#0000ff">01991_16_00069</span>
        </a>
        <span style="font-size:9pt">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 9pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span>Cordialement / Kind Regards,</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 10pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-family:Verdana; font-weight:bold; color:#0c64c0">{ca_name} – SC62420</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 8pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-family:Verdana; font-style:italic; color:#ff0000">ENG/EES/EEEM/EEWE/EAWE/AIWE/EAIW</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 8pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-family:Verdana; font-style:italic; color:#002451">EE Architect Cockpit, HMI and Connected Services</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 13.5pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-style:italic">--------------------------------------------------------------------------------------</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 8pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-family:Verdana; font-style:italic; color:#5e327c">Consultant du Groupe Capgemini Engineering pour STELLANTIS</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;font-size: 8pt;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <span style="font-family:Verdana; font-style:italic; color:#5e327c">Casanearshore, Shore 17, 5ème étage, 20270, CASABLANCA, Maroc</span>
      </p>
      <p style="margin-bottom: 0pt;line-height: normal;background-color: #ffffff;margin: 0pt 0pt 8pt;">
        <img src="images/Aspose.Words.09e1a2d9-057e-410d-8338-4fd164e6627c.001.png" width="487" height="50" alt="" style="-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline">
      </p>
      <p style="margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
      <p style="margin: 0pt 0pt 8pt;">
        <span style="-aw-import:ignore">&#xa0;</span>
      </p>
    </div>
  </body>
</html>

        '''

        email_file_path = 'C:/Users/nizar/Application_cap/EMAIL'  # Update with your desired path
        email_file_name = f'{ca_name}_{function_name}_email.html'
        full_email_path = os.path.join(email_file_path, email_file_name)

        # Save the email content to the defined path
        with open(full_email_path, 'w', encoding='utf-8') as file:
            file.write(email_body)

        msg.add_alternative(email_body, subtype='html')
        msg['From'] = 'your_email@gmail.com'
        msg['To'] = recipient_email

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)

        confirmation_label.config(text='Email sent successfully.')
    else:
        confirmation_label.config(text='Please enter a valid email address.')

    return email_body


root = tk.Tk()
root.title('EEA Data Management')

main_frame = ttk.Frame(root, padding='20')
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

choose_csv_button = ttk.Button(main_frame, text='Choose CSV File', command=choose_csv_file)
choose_csv_button.grid(row=0, column=0, pady=10, padx=10)

ca_combobox_label = ttk.Label(main_frame, text='Select CA:')
ca_combobox_label.grid(row=1, column=0, pady=5, padx=5, sticky=tk.W)

ca_combobox = ttk.Combobox(main_frame, state='readonly')
ca_combobox.grid(row=1, column=1, pady=5, padx=5, sticky=tk.W)

function_combobox_label = ttk.Label(main_frame, text='Select Function:')
function_combobox_label.grid(row=2, column=0, pady=5, padx=5, sticky=tk.W)

function_combobox = ttk.Combobox(main_frame, state='readonly')
function_combobox.grid(row=2, column=1, pady=5, padx=5, sticky=tk.W)

ca_combobox.bind("<<ComboboxSelected>>", show_functions)
'''
recipient_label = ttk.Label(main_frame, text='Recipient Email:')
recipient_label.grid(row=14, column=0, pady=5, padx=5, sticky=tk.W)

recipient_entry = ttk.Entry(main_frame)
recipient_entry.grid(row=14, column=1, pady=5, padx=5, sticky=tk.W)
'''
end_date_label = ttk.Label(main_frame, text='End Date (YYYY-MM-DD):')
end_date_label.grid(row=15, column=0, pady=5, padx=5, sticky=tk.W)

end_date_entry = ttk.Entry(main_frame)
end_date_entry.grid(row=15, column=1, pady=5, padx=5, sticky=tk.W)

ee_arch_label_cr = ttk.Label(main_frame, text='EE Architecture Check Request:')
ee_arch_label_cr.grid(row=3, column=0, pady=5, padx=5, sticky=tk.W)

ee_arch_entry_cr = ttk.Combobox(main_frame, values=["Yes", "No"])
ee_arch_entry_cr.grid(row=3, column=1, pady=5, padx=5, sticky=tk.W)

'''
ee_arch_label_team = ttk.Label(main_frame, text='EE Architecture Team:')
ee_arch_label_team.grid(row=4, column=0, pady=5, padx=5, sticky=tk.W)

ee_arch_entry_team = ttk.Entry(main_frame)
ee_arch_entry_team.grid(row=4, column=1, pady=5, padx=5, sticky=tk.W)
'''

ee_safety_label_cr = ttk.Label(main_frame, text='EE Safety Check Request:')
ee_safety_label_cr.grid(row=5, column=0, pady=5, padx=5, sticky=tk.W)

ee_safety_entry_cr = ttk.Combobox(main_frame, values=["Yes", "No"])
ee_safety_entry_cr.grid(row=5, column=1, pady=5, padx=5, sticky=tk.W)

'''
ee_safety_label_team = ttk.Label(main_frame, text='EE Safety Team:')
ee_safety_label_team.grid(row=6, column=0, pady=5, padx=5, sticky=tk.W)

ee_safety_entry_team = ttk.Entry(main_frame)
ee_safety_entry_team.grid(row=6, column=1, pady=5, padx=5, sticky=tk.W)
'''

ee_diagnosis_label_cr = ttk.Label(main_frame, text='EE Diagnosis Check Request:')
ee_diagnosis_label_cr.grid(row=7, column=0, pady=5, padx=5, sticky=tk.W)

ee_diagnosis_entry_cr = ttk.Combobox(main_frame, values=["Yes", "No"])
ee_diagnosis_entry_cr.grid(row=7, column=1, pady=5, padx=5, sticky=tk.W)

'''
ee_diagnosis_label_team = ttk.Label(main_frame, text='EE Diagnosis Team:')
ee_diagnosis_label_team.grid(row=8, column=0, pady=5, padx=5, sticky=tk.W)

ee_diagnosis_entry_team = ttk.Entry(main_frame)
ee_diagnosis_entry_team.grid(row=8, column=1, pady=5, padx=5, sticky=tk.W)
'''

tech_review_label_cr = ttk.Label(main_frame, text='Technical Review Check Request:')
tech_review_label_cr.grid(row=9, column=0, pady=5, padx=5, sticky=tk.W)

tech_review_entry_cr = ttk.Combobox(main_frame, values=["Yes", "No"])
tech_review_entry_cr.grid(row=9, column=1, pady=5, padx=5, sticky=tk.W)

'''
tech_review_label_team = ttk.Label(main_frame, text='Technical Review Team:')
tech_review_label_team.grid(row=10, column=0, pady=5, padx=5, sticky=tk.W)

tech_review_entry_team = ttk.Entry(main_frame)
tech_review_entry_team.grid(row=10, column=1, pady=5, padx=5, sticky=tk.W)
'''

show_details_button = ttk.Button(main_frame, text='Show Details', command=show_details)
show_details_button.grid(row=11, column=0, pady=10, padx=10, sticky=tk.W)

send_email_button = ttk.Button(main_frame, text='Save Email', command=send_email)
send_email_button.grid(row=11, column=1, pady=10, padx=10, sticky=tk.W)

details_text = tk.Text(main_frame, wrap=tk.WORD, state=tk.DISABLED, width=50, height=10)
details_text.grid(row=12, column=0, columnspan=2, padx=10, pady=10, sticky=tk.W)

confirmation_label = ttk.Label(main_frame, text='')
confirmation_label.grid(row=13, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)

root.mainloop()
