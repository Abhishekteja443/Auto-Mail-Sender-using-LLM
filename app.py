from ollama import Client
import ollama
import streamlit as st
import re
import csv
from csv import writer
from datetime import datetime
import os
import pythoncom
import logging
import win32com.client as win32
import uuid  # Import the uuid module

# Initialize logging
logging.basicConfig(level=logging.INFO, filename="app.log", 
                    filemode="a", format="%(asctime)s - %(levelname)s - %(message)s")

# Initialize the Ollama client
client = Client(host='http://localhost:11434')

# Define the model to use
model = 'llama3.2:3b'  # Ensure the model name is correct

# Function to validate email
def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(pattern, email)


def is_email_in_logs(email):
    """Check if the email already exists in logs.csv"""
    try:
        with open('logs.csv', 'r', newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if row and row[2] == email:  # Assuming email is in the 3rd column
                    return True
    except FileNotFoundError:
        return False  # If the file does not exist, no emails are logged yet
    return False

# Function to send email
def send_email(subject, body, recipient, attachment):
    try:
        pythoncom.CoInitialize()  # Initialize COM
        outlook = win32.Dispatch('outlook.application')
        olMailItem = 0x0

        # body = body.replace('\\n', '\n')

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
        logging.info(f"Email sent successfully to {recipient}!")
        return True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM after operations

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
        index_s=full_response.find('{')
        full_response=full_response[index_s+1:]
        
        s_start=full_response.find("subject")
        subject=full_response[s_start+11:full_response.find('",')]
        b_start=full_response.find("body")
        body=full_response[b_start+8:full_response.find('}')-2] 
        logging.info("Email generated successfully to.")
        return subject,body
    except Exception as e:
        st.error(f"An error occurred: {e}")
        logging.error(f"Error generating email: {e}")
        return None,None


# Main function
def main():
    st.title("Auto Email Sender for Job Applications")
    
    # Input fields
    prof_name = st.text_area("Enter professor name")
    prof_mail = st.text_area("Enter professor mail id")
    job_description = st.text_area("Enter job Description")
    
    # Validate email
    if prof_mail and not is_valid_email(prof_mail):
        st.error("Invalid email address format.")
        return
    
    # Job type selection
    job_type = st.radio("Job Type", ("Technical", "Non-Technical"))
    
    # Check if all fields are filled
    if not prof_name or not prof_mail or not job_description:
        st.warning("Please fill in all required fields.")
        return
    
    if is_email_in_logs(prof_mail):
        st.warning("An email has already been sent to this recipient. Email generation is disabled.")
        return
    
    # Define system prompt based on job type
    if job_type == "Technical":
        resume = "Your resume Url"
        about_me = "My name is Abhishek Teja Goli, and I am a graduate student pursuing my Master’s in Data Science at Wright State University. With a strong foundation in Python, machine learning, and data analysis. My academic training and internship experience have equipped me with technical expertise, while my problem-solving skills and ability to explain complex concepts make me an effective mentor and collaborator. I am passionate about supporting academic and research initiatives by assisting students, conducting data-driven research, and streamlining analytical processes. You can reach me at goli.34@wright.edu."

        system_prompt = """
            You are an advanced assistant designed to craft professional emails. Based on the provided context, your task is to generate a formal email that aligns with the tone and structure of the example below. Ensure the email includes a proper subject line, a respectful salutation, a concise and clear body, and a polite closing with appropriate contact information. The output should be humanly written and must not contain any placeholders like [insert data] or [ ].

            The email should clearly mention that the user is looking for on-campus job opportunities under the professors mentioned, which could include roles for assistants positions.
            
            Return the response in JSON format with two keys:
            - `"subject"`: The subject of the email.
            - `"body"`: The body of the email.

            The output should strictly follow this format:

            {
            "subject": "[Insert subject line based on context]",
            "body": "Dear [Recipient's Name/Title],

            [Opening line introducing yourself and providing context.]

            [Main body of the email, explaining your purpose, relevant qualifications, and any specific requests. Keep the tone formal and professional.]

            [Closing line expressing gratitude and anticipation of a response.]

            Best regards,  
            [Your Name]  
            [Your Email Address] "
            }
        """
    else:
        resume = "Your resume Url"
        about_me = "My name is Abhishek Teja Goli, and I am a graduate student pursuing my Master’s in Data Science at Wright State University. With strong communication, teamwork, and time management skills, I am eager to contribute to campus services . My experience in customer service, tutoring, and collaborative projects has strengthened my ability to engage with diverse individuals, handle responsibilities efficiently, and provide excellent assistance. I am passionate about creating a welcoming and organized environment where students and faculty can have a positive experience. You can reach me at goli.34@wright.edu."

        system_prompt = """
            You are an advanced assistant designed to craft professional emails based on the provided context. Your task is to generate a formal email that aligns with the tone and structure of the example below. Ensure the email includes a proper subject line, a respectful salutation, a concise and clear body, and a polite closing with appropriate contact information. The output should be humanly written and must not contain any placeholders like [insert data] or [ ].

            Additionally, the email should focus on non-technical positions. You should tailor the content to reflect a genuine interest in these types of positions, emphasizing skills such as communication, organization, and teamwork.
            
            Return the response in JSON format with two keys:
            - `"subject"`: The subject of the email.
            - `"body"`: The body of the email.

            The output should strictly follow this format:

            {
            "subject": "[Insert subject line based on context]",
            "body": "Dear [Recipient's Name/Title],

            [Opening line introducing yourself and providing context.]

            [Main body of the email, explaining your purpose, relevant qualifications, and any specific requests. Keep the tone formal and professional.]

            [Closing line expressing gratitude and anticipation of a response.]

            Best regards,  
            [Your Name]  
            [Your Email Address] "
            }
        """

    question = f'Write a mail to {prof_name}, {prof_mail}. {job_description}, {about_me}.'

    def check_body(body):
            body=body.strip()
            if body[-1]=='d':
                body=body+'u'
            elif body[-1]=='e':
                body=body+'du'
            elif body[-1]==".":
                body=body+'edu'
            return body

    # Generate email button
    if st.button("Generate Email"):
        subject, body = generate_email(system_prompt, question)
        # while '\n' in body:
        #     subject, body = generate_email(system_prompt, question)
        body=check_body(body)
        if subject and body:
            st.session_state["subject"] = subject
            st.session_state["body"] = body
            st.success("Email generated successfully!")
        else:
            st.error("Failed to generate email.")

    # Display the generated email if available
    if "subject" in st.session_state and "body" in st.session_state:
        st.subheader("Generated Email")
        # st.text_area("Subject", , height=300)
        st.text_area("Body",st.session_state["subject"]+ st.session_state["body"], height=300)

        # Regenerate email button
        if st.button("Regenerate Email"):
            subject, body = generate_email(system_prompt, question)
            # while '\n' in body:
            #     subject, body = generate_email(system_prompt, question)
            body=check_body(body)
            if subject and body:
                st.session_state["subject"] = subject
                st.session_state["body"] = body
                st.success("Email regenerated successfully!")
            else:
                st.error("Failed to regenerate email.")
            st.text_area("Body",st.session_state["subject"]+ st.session_state["body"], height=300)

        # Send email button
        if st.button("Send Email"):
            with st.spinner("Sending email..."):
                if send_email(st.session_state["subject"], st.session_state["body"], prof_mail, resume):
                    # Generate a unique ID for the log entry
                    unique_id = str(uuid.uuid4())
                    
                    # Log the data with the unique ID
                    with open('logs.csv', 'a', newline='') as f_object:
                        now = datetime.now()
                        current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")
                        writer_object = writer(f_object)
                        writer_object.writerow([unique_id, prof_name, prof_mail, job_type, job_description, 1, current_datetime])
                    st.success("Email and data stored successfully!")
                else:
                    st.error("Failed to send email.")

# Run the app
if __name__ == "__main__":
    main()
