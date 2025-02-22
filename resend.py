import csv
from datetime import datetime, timedelta
import os
import pythoncom
import logging
import win32com.client as win32
from ollama import Client
import ollama
import gc  # Add garbage collection
import re
# Initialize logging
logging.basicConfig(level=logging.INFO, filename="resend.log", 
                    filemode="a", format="%(asctime)s - %(levelname)s - %(message)s")

# Initialize the Ollama client
client = Client(host='http://localhost:11434')

# Define the model to use
model = 'llama3.2:3b'  # Ensure the model name is correct

# Function to send email


def send_email(subject, body, recipient, attachment):
    try:
        pythoncom.CoInitialize()  # Initialize COM
        outlook = win32.Dispatch('outlook.application')
        olMailItem = 0x0
        print(recipient)
        print(body)
        print()
        body=re.sub(r"\\n", "\n", body)
        # print("*****")

        # body=body.replace('\n\n','\n\n')
        # body = body.replace('\\n', '\n')
        # body=bodreplace('\n','\n')
        

        newMail = outlook.CreateItem(olMailItem)
        newMail.Subject = subject
        newMail.BodyFormat = 1  # olFormatPlain
        newMail.Body = body  # Plain text body
        newMail.To = recipient  # Single recipient

        logging.info("Data loaded")
        if os.path.exists(attachment):
            newMail.Attachments.Add(attachment)
            logging.info("Document loaded")
        else:
            logging.error(f"Attachment not found: {attachment}")

        newMail.Send()  # Sends the email immediately
        logging.info("Email sent successfully!")
        return True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return False
    finally:
        newMail = None  # Release the email object
        outlook = None  # Release Outlook instance
        pythoncom.CoUninitialize()  # Uninitialize COM
        gc.collect()  # Run garbage collection to clean up COM objects

# Function to generate email using Ollama
def generate_email(system_prompt, question):
    try:
        response = ollama.chat(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": question},
            ],
        )
        full_response = response.get("message", {}).get("content", "No response generated")
        
        # Parse the JSON response
        index_s = full_response.find('{')
        full_response = full_response[index_s + 1:]
        
        s_start = full_response.find("subject")
        subject = full_response[s_start + 11:full_response.find('",')]
        b_start = full_response.find("body")
        body = full_response[b_start + 8:full_response.find('}') - 2]
        logging.info("Email generated successfully.") 
        return subject, body
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None, None
        

# Function to resend email
def resend_email(prof_name, prof_mail, job_description, status, job_type,time_sent):    
    # Define system prompt based on job type and status
    if status >= 1:
        if job_type == "Technical":    
                about_me = "My name is Abhishek Teja Goli, and I am a graduate student pursuing my Master’s in Data Science at Wright State University. With a strong foundation in Python, machine learning, and data analysis, I am eager to contribute as a Teaching Assitant. My academic training and internship experience have equipped me with technical expertise, while my problem-solving skills and ability to explain complex concepts make me an effective mentor and collaborator. I am passionate about supporting academic and research initiatives by assisting students, conducting data-driven research, and streamlining analytical processes. You can reach me at goli.34@wright.edu."

                system_prompt = """
                    You are an advanced assistant designed to craft professional follow-up emails for technical job opportunities. You should tailor the content to reflect a genuine interest in these types of positions, emphasizing skills related to computer and software. Based on the provided context, your task is to generate a formal follow-up email that aligns with the tone and structure of the example below. Ensure the email includes a proper subject line, a respectful salutation, a concise and clear body, and a polite closing with appropriate contact information. 
                    
                    Note: The output should be humanly written with 3 paraghraphs only.

                    The email should clearly mention that the user is following up on their previous email regarding on-campus technical job opportunities under the professors mentioned, which could include roles for assistants positions with relevent skills.
                    
                    Return the response in JSON format with two keys:
                    
                    - `"subject"`: The subject of the email.
                    - `"body"`: The body of the email.

                    The output should strictly follow this format:

                    {
                    "subject": "[Insert subject line based on context]",
                    "body": "Dear [Recipient's Name/Title],

                    [Opening line introducing yourself and reminding them of your previous email.]

                    [Main body of the email, explaining your purpose, relevant technical qualifications, and any specific requests. Keep the tone formal and professional.]

                    [Closing line expressing gratitude and anticipation of a response.]

                    Best regards,  
                    [Your Name]  
                    [Your Email Address]"
                    }
                """
        
        else:  # Non-Technical
                about_me = "My name is Abhishek Teja Goli, and I am a graduate student pursuing my Master’s in Data Science at Wright State University. With strong communication, teamwork, and time management skills, I am eager to contribute to campus services such as recreation, dining, or student support roles. My experience in customer service, tutoring, and collaborative projects has strengthened my ability to engage with diverse individuals, handle responsibilities efficiently, and provide excellent assistance. I am passionate about creating a welcoming and organized environment where students and faculty can have a positive experience. You can reach me at goli.34@wright.edu."

                system_prompt = """
                    You are an advanced assistant designed to craft professional follow-up emails for non-technical job opportunities. You should tailor the content to reflect a genuine interest in these types of positions, emphasizing skills such as communication, organization, and teamwork. Based on the provided context, your task is to generate a formal follow-up email that aligns with the tone and structure of the example below. Ensure the email includes a proper subject line, a respectful salutation, a concise and clear body, and a polite closing with appropriate contact information. 
                    
                    Note: The output should be humanly written with 3 paraghraphs only.

                    The email should clearly mention that the user is following up on their previous email regarding non-techincal on-campus job opportunities under the professors mentioned.
                    
                    Return the response in JSON format with two keys:
                    - `"subject"`: The subject of the email.
                    - `"body"`: The body of the email.

                    The output should strictly follow this format:

                    {
                    "subject": "[Insert subject line based on context]",
                    "body": "Dear [Recipient's Name/Title],

                    [Opening line introducing yourself and reminding them of your previous email.]

                    [Main body of the email, explaining your purpose, relevant non-technical qualifications, and any specific requests. Keep the tone formal and professional.]

                    [Closing line expressing gratitude and anticipation of a response.]

                    Best regards,  
                    [Your Name]  
                    [Your Email Address]"
                    }
                """
    else:
        return None, None  # Do not send email if status is negative

    question = f'Write a follow-up mail to {prof_name}, {prof_mail}. {job_description}, {about_me}, previous mail sent date : {time_sent}.'
    subject, body = generate_email(system_prompt, question)
    while '[' in subject or '[' in body:
        subject,body=generate_email(system_prompt,question)
        logging.info("Mail regenrated!!")
    return subject, body

# Function to update logs.csv
def update_logs(unique_id, status, new_time):
    rows = []
    with open('C:/Users/gabhi/Projects/Auto_mail_sender/logs.csv', 'r') as file:
        reader = csv.reader(file)
        rows = list(reader)

    for row in rows:
        if row[0] == unique_id:
            row[5] = status  # Update status
            row[6] = new_time  # Update time

    with open('C:/Users/gabhi/Projects/Auto_mail_sender/logs.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(rows)

# Main function to process logs and resend emails
def process_logs():
    with open('C:/Users/gabhi/Projects/Auto_mail_sender/logs.csv', 'r') as file:
        reader = csv.reader(file)
        logs = list(reader)

    # Skip the header if present
    if logs[0][0].lower() == "uid":  
        logs = logs[1:]  

    for row in logs:
        unique_id, prof_name, prof_mail, job_type, job_description, status, time_str = row
        # print(time_str)
        
        # Ensure time_str is not empty or invalid
        if time_str.lower() == "time" or not time_str.strip():
            logging.error(f"Skipping row with invalid time: {row}")
            continue

        try:
            time_sent = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except ValueError as e:
            logging.error(f"Error parsing time '{time_str}': {e}")
            continue
        
        def check_body(body):
            body=body.strip()
            if body[-1]=='d':
                body=body+'u'
            elif body[-1]=='e':
                body=body+'du'
            elif body[-1]==".":
                body=body+'edu'
            return body
            
        current_time = datetime.now()
        if current_time - time_sent >= timedelta(weeks=1):
            if int(status) >= 0:
                subject, body = resend_email(prof_name, prof_mail, job_description, int(status), job_type,time_sent)
                # while "\n" in body:
                #     subject, body = resend_email(prof_name, prof_mail, job_description, int(status), job_type,time_sent)
                if "\n" in body:
                    print("naiuwehfuhvv")
                body=check_body(body)
                # print(subject,body)
                if subject and body:
                    if job_type=="Technical":
                        resume = "C:/Users/gabhi/Projects/Auto_mail_sender/Abhishek Teja Goli-Resume.pdf"
                    else:
                        resume = "C:/Users/gabhi/Projects/Auto_mail_sender/Abhishek_Teja_Goli-Resume.pdf"
                    if send_email(subject, body, prof_mail, resume):
                        new_status = int(status) + 1
                        new_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        update_logs(unique_id, new_status, new_time)
                        logging.info(f"Email resent to {prof_name} ({prof_mail}) with status {new_status}.")
                    else:
                        logging.error(f"Failed to resend email to {prof_name} ({prof_mail}).")
                else:
                    logging.error(f"Failed to generate email for {prof_name} ({prof_mail}).")


# Run the process
if __name__ == "__main__":
    process_logs()