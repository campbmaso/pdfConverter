import PyPDF2
# from docx import Document
# from docx.shared import Pt
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
import boto3
import openai
from openai import OpenAI

from copy import deepcopy
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import os
import base64
import requests
import time
import json


# AWS Secrets Manager client
dynamodb = boto3.resource('dynamodb', region_name='us-east-2')
api_usage_table = dynamodb.Table('APIKeyUsage')
client = boto3.client(service_name='secretsmanager', region_name="us-east-2")
SECRET_NAME = "openai_secrets"

def get_secret_value(api_key_identifier):

    try:
        get_secret_value_response = client.get_secret_value(SecretId=SECRET_NAME)
    except Exception as e:
        # Handle exceptions as necessary.
        print(f"Exception: {e}")
        return None
    
    # Extract the secret JSON
    secret_values = eval(get_secret_value_response['SecretString'])  # Convert string to dictionary
    
    # Return the specific API key for the provided identifier.
    return secret_values.get(api_key_identifier)

def select_api_key():
    # Fetch all items (API keys and their counts) from DynamoDB.
    response = api_usage_table.scan()
    keys_data = response['Items']
    
    # If the table is empty, initialize it with the API key identifiers and counts of 0.
    if not keys_data:
        api_key_identifiers = [f"api_key{i}" for i in range(1, 26)]
        for key_id in api_key_identifiers:
            api_usage_table.put_item(Item={'APIKey': key_id, 'count': 0}) 
        
        # Re-scan to get the initialized data.
        response = api_usage_table.scan()
        keys_data = response['Items']

    # Sort the items based on the count (ascending) and select the first item (least used key).
    selected_key_data = sorted(keys_data, key=lambda x: x['count'])[0]
    
    # Increment the count for the selected key.
    api_usage_table.update_item(
        Key={'APIKey': selected_key_data['APIKey']},
        UpdateExpression='SET #count_attribute = #count_attribute + :increment',
        ExpressionAttributeNames={'#count_attribute': 'count'},
        ExpressionAttributeValues={':increment': 1}
    )

    
    # Return the API key identifier (like "api_key1").
    return selected_key_data['APIKey']

selected_key_identifier = select_api_key()
actual_api_key_value = get_secret_value(selected_key_identifier)
client = OpenAI(api_key=actual_api_key_value)


s3_client = boto3.client("s3")
BUCKET_NAME = os.environ.get("resume_bucket")

# open the file: "/Users/campbmaso/Desktop/Development/GitHub/Sandbox/Backend/resumes/TEMPLATE - Pscyence Resume Service (1).docx"

# =========================== get resume files ready =============================================================================================

resume_text = ""
def get_resume_text(filename):
    resume_text = ""
    s3_client = boto3.client("s3")
    response = s3_client.get_object(Bucket="resume-s3bucket", Key=filename)
    # Read the PDF file content from S3 directly into a BytesIO object
    file_content = BytesIO(response["Body"].read())
    reader = PyPDF2.PdfReader(file_content)
    for page in reader.pages:
        resume_text += page.extract_text()
    print(f"full resume text: {resume_text}")
    return resume_text


# =========================== data templates =============================================================================================
dummy_resume = """
1
Cory G. Mazure49225 Equestrian Court, Chesterfield, MI 48047
Cell: (586) 961-5274 cory.mazure@gmail.com
LinkedIn: www.linkedin.com/in/cory-mazure
Work Experience
Process Engineer, Golf Ball R&D
Callaway Golf, Carlsbad, CA (January 2023 - Present)
● Lead engineer behind the experimentation, validation, and implementation of new spray gun equipment
in production for 3 paint lines, resulting in a more controllable process and therefore quality product
● Organized DOEs to study differences in mechanical processes between R&D and production, reducing
variance in processes, ultimately establishing confidence which accelerated new ball development
● Implemented new test methods to validate chemistry, leading to new proprietary golf ball features
● Utilized injection molding modeling software to modify tooling, improving concentricity and yield
Mechanical Engineering, Vehicle Safety Engineering
Stellantis, Auburn Hills, MI (Summer of 2019 – Summer 2022)
● Led a design for six sigma project for optimizing the amount of test dummies in use, saving $250K/year
● Reduced vehicle validation timing by developing test fixtures for high inertia door latch testing
● Analyzed channel data for 80+ test rib deflection study, resulting in global program airbag innovations
● Coordinated JD Power quality study and benchmarking for heated steering wheel (HSW) optimization
● Strengthened best practices for HSW calibration ranges, directly applied within 2 new flagship
programs, resulting in enhanced customer satisfaction
● Investigated 2 competing suppliers for hands on detection mats via Minitab statistical analysis, leading
to more competitive pricing
● Established statistical analysis of the dummy labs 4 certification machines, reducing certification timing
and associated costs by $20K/year
● Validated 2 energy absorbing knee brackets and a driver airbag, achieving higher safety ratings
Leadership Experience
● Kettering Student Government Operations Council – Director: Planned and executed over 15 events a
term, communicating progress of the council weekly to other government sub-sections
● Sigma Chi Distant Leadership Course: Learned practical strategies for remote team communication
● Golf Club – President: Ran weekly meetings with passion, and liaised contracts with local driving range
● Mindfulness Club – President: Facilitated awareness of mental health practices and built a community
● Inter Fraternity Council – Philanthropy Chair: Orchestrated 2 events to improve the image of campus
Skills
● Stellantis Sponsored Courses: Attended 5 in depth training courses for DFMEA, DFSS, and GD&T
● Moldflow injection molding simulation software
● Proficient in NX, Fusion 360, & SolidWorks: CAE and FEA
● Experienced in Minitab statistical analysis software for interpreting DOE results
● Utilized Altair Hyperworks to analyze vehicle crash data
● Avid user of Microsoft Office and Google suite for statistical analysis and project management
Education
Bachelor of Science in Mechanical Engineering
Kettering University Flint, MI (Oct. 2019 - Dec. 2022)
● Kettering University Honors Program: Completing 2 advanced projects per academic term
● Thesis Project: WorldSID ATD Rib Deflection and Seating Analysis Assessment
● GPA: 3.95/4.00, Summa Cum Laude, Dean's List every academic term
"""

json_template = {
    "personal_information": {
        "name": "John Doe",
        "email": "john.doe@example.com",
        "phone": "123-456-7890",
        "linkedin": "linkedin.com/in/johndoe",
    },
    "summary": "Experienced software engineer with a strong background in developing scalable applications and improving software development processes.",
    "skills": [
        "Python",
        "Java",
        "Docker",
        "Kubernetes",
        "React",
        "Agile methodologies",
    ],
    "additional_activities": [
        {
            "title": "Hackathon Participant",
            "bullets": [
                "Participated in a 24-hour hackathon, developing a web application that won 2nd place."
            ],
            "date": "March 2020",
        },
    ],
    "work_experience": [
        {
            "job_title": "Senior Software Engineer",
            "company": "Tech Solutions Inc.",
            "location": "San Francisco, CA",
            "start_date": "June 2018",
            "end_date": "Present",
            "bullets": [
                "Led the development of a new feature that increased user engagement by 20%",
                "Implemented a CI/CD pipeline that reduced deployment time by 50%",
                "Mentored junior developers, improving team productivity and knowledge sharing",
            ],
        },
        {
            "job_title": "Software Engineer",
            "company": "Innovate Startup",
            "location": "San Francisco, CA",
            "date_range": "June 2016 - June 2018",
            "bullets": [
                "Developed a high-traffic web application using React and Node.js",
                "Optimized database queries, reducing load times by 30%",
                "Collaborated with cross-functional teams to define, design, and ship new features",
            ],
        },
    ],
    "education": [
        {
            "degree": "Bachelor of Science in Computer Science",
            "institution": "University of Example",
            "location": "Example City",
            "date_range": "September 2010 - June 2014",
            "bullets": "Graduated with honors",
        }
    ],
    "certifications": [
        {
            "title": "Certified Kubernetes Administrator",
            "issuer": "The Linux Foundation",
            "date_issued": "July 2019",
        }
    ],
    "projects": [
        {
            "title": "Personal Portfolio Website",
            "bullets": [
                "A personal website to showcase my projects and resume. Built with HTML, CSS, and JavaScript."
            ],
            "link": "http://www.johndoeportfolio.com",
        }
    ],
}

header_template = {
    "personal_information": {
        "name": "John Doe",
        "email": "john.doe@example.com",
        "phone": "123-456-7890",
        "linkedin": "linkedin.com/in/johndoe",
    }
}
summary_template = {
    "summary": "Experienced software engineer with a strong background in developing scalable applications and improving software development processes."
}
work_experience_template = {
    "work_experience": [
        {
            "job_title": "Senior Software Engineer",
            "company": "Tech Solutions Inc.",
            "location": "San Francisco, CA",
            "date_range": "June 2018 - Present",
            "bullets": [
                "Led the development of a new feature that increased user engagement by 20%",
                "Implemented a CI/CD pipeline that reduced deployment time by 50%",
                "Mentored junior developers, improving team productivity and knowledge sharing",
            ],
        },
        {
            "job_title": "Software Engineer",
            "company": "Innovate Startup",
            "location": "San Francisco, CA",
            "date_range": "June 2016 - June 2018",
            "bullets": [
                "Developed a high-traffic web application using React and Node.js",
                "Optimized database queries, reducing load times by 30%",
                "Collaborated with cross-functional teams to define, design, and ship new features",
            ],
        },
    ],
}
additional_activities_template = (
    {
        "additional_activities": [
            {
                "title": "Hackathon Participant",
                "bullets": [
                    "Participated in a 24-hour hackathon, developing a web application that won 2nd place."
                ],
                "date": "March 2020",
            },
        ],
        "leadership": [
            {
                "title": "Director, Student Government Operations Council",
                "bullets": [
                    "Planned and executed over 15 events a term, communicating progress of the council weekly to other government sub-sections",
                    "Led a team of 20 students in planning and executing events",
                ],
                "date": "March 2020",
            },
        ],
    },
)
skills_template = {
    "skills": [
        "Python",
        "Java",
        "Docker",
        "Kubernetes",
        "React",
        "Agile methodologies",
    ]
}
education_template = {
    "education": [
        {
            "degree": "Bachelor of Science in Computer Science",
            "institution": "University of Example",
            "location": "Example City",
            "date_range": "September 2010 - June 2014",
            "bullets": "Graduated with honors",
        },
        {
            "degree": "Certified Kubernetes Administrator",
            "institution": "The Linux Foundation",
            "date_range": "July 2019",
        },
    ],
}
certifications_template = {
    "certifications": [
        {
            "title": "Certified Kubernetes Administrator",
            "issuer": "The Linux Foundation",
            "date_issued": "July 2019",
        }
    ],
}

# =========================== get resume sections =============================================================================================


def get_header_section(resume):
    start = time.time()
    resume = " ".join(resume.split()[:50])
    print(f"SPLIT resume: {resume}")
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user header data, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"Please find the user header data and return it in JSON format like this example: {header_template}. \n Here is the resume: {resume}.",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No header found.")
        return None
    print(f"response: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def get_summary_section(resume):
    start = time.time()
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user summary statement data, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"""Please find the user summary/objective statement and return it in JSON format like this example: {summary_template}. \n Here is the resume: {resume}. 
                If the user does not have a summary statement already present in their resume, please return an empty string.""",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No summary found.")
        return None
    print(f"response: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def get_work_experience_section(resume):
    start = time.time()
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user work experiences from their resume, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"""Please find the user work experiences from their resume and return it in JSON format like this example: {work_experience_template}. \n Here is the resume: {resume}. 
                If the user does not have an work experience already present in their resume, please return an empty string.""",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No work experience found.")
        return None
    print(f"response for WE section: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def get_additional_activities_section(resume):
    start = time.time()
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user additional activities or miscellaneous section from their resume, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"""Please find the user additional activities or miscellaneous section from their resume and return it in JSON format like this example: {additional_activities_template}. \n Here is the resume: {resume}. 
                If the user does not have an additional activities or miscellaneous section already present in their resume, please return an empty string.""",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No additional activities found.")
        return None
    print(f"response for AA / Le section: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def get_skills_section(resume):
    start = time.time()
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user skills section from their resume, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"""Please find the user skills section from their resume and return it in JSON format like this example: {json_template}. \n Here is the resume: {resume}. 
                If the user does not have a skills section already present in their resume, please return an empty string.""",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No additional activities found.")
        return None
    print(f"response skills section: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def get_education_section(resume):
    start = time.time()
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user education section from their resume, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"""Please find the user skills section from their resume and return it in JSON format like this example: {education_template}. \n Here is the resume: {resume}. 
                If the user does not have a education section already present in their resume, please return an empty string.""",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No additional activities found.")
        return None
    print(f"response EDU: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def get_certifications_section(resume):
    start = time.time()
    completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"Please find the user certifications section from their resume, and return it in JSON format.",
            },
            {
                "role": "user",
                "content": f"""Please find the user certifications section from their resume and return it in JSON format like this example: {certifications_template}. \n Here is the resume: {resume}. 
                If the user does not have a certifications section already present in their resume, please return an empty string.""",
            },
        ],
        temperature=0.2,
    )
    section = completion.choices[0].message.content
    if section == "":
        print("No additional activities found.")
        return None
    print(f"response: {section}")
    print(f"GPT time took: {time.time() - start} seconds.\n")
    return section


def convert_string_to_json(string):
    try:
        json_formatted = json.loads(string)
        print(f"converted JSON: {json_formatted} from string: {string}")
        return json_formatted
    except json.JSONDecodeError as e:
        print(f"Error: {e}")
        return None


def generate_sections(dummy_resume, concurrent=True):
    if concurrent == False:
        # ============================ Call the functions in serial to parse the sections ============================
        user_header_data = convert_string_to_json(get_header_section(dummy_resume))
        user_summary_data = convert_string_to_json(get_summary_section(dummy_resume))
        work_experience_data = convert_string_to_json(get_work_experience_section(dummy_resume))
        additional_activities_data = convert_string_to_json(get_additional_activities_section(dummy_resume))
        skills_data = convert_string_to_json(get_skills_section(dummy_resume))
        education_data = convert_string_to_json(get_education_section(dummy_resume))
        certifications_data = convert_string_to_json(get_certifications_section(dummy_resume))
    else:
        # ============================ Call the functions concurrently to parse the sections ============================
        with ThreadPoolExecutor(max_workers=7) as executor:
            future_header = executor.submit(get_header_section, dummy_resume)
            future_summary = executor.submit(get_summary_section, dummy_resume)
            future_work_experience = executor.submit(
                get_work_experience_section, dummy_resume
            )
            future_additional_activities = executor.submit(
                get_additional_activities_section, dummy_resume
            )
            future_skills = executor.submit(get_skills_section, dummy_resume)
            future_education = executor.submit(get_education_section, dummy_resume)
            # future_certifications = executor.submit(get_certifications_section, dummy_resume)

            user_header_data = convert_string_to_json(future_header.result())
            user_summary_data = convert_string_to_json(future_summary.result())
            work_experience_data = convert_string_to_json(future_work_experience.result())
            additional_activities_data = convert_string_to_json(future_additional_activities.result())
            skills_data = convert_string_to_json(future_skills.result())
            education_data = convert_string_to_json(future_education.result())
            # certifications_data = convert_string_to_json(future_certifications.result())

    # combine the parsed sections into a single JSON object
    parsed_user_data = {
        **user_header_data,
        **user_summary_data,
        **work_experience_data,
        **additional_activities_data,
        **skills_data,
        **education_data,
        # **certifications_data,
    }
    print(f"parsed user data type: {type(parsed_user_data)}")
    return parsed_user_data


# ===================================================== Function executions =================================================================

def lambda_handler(event, context):
    print(f"event: {event}")
    execution_start = time.time()
    event_body = event.get('body')
    
    if event_body.get("local_testing"):
        filename = event_body.get('filename')
        resume_text = get_resume_text(filename) # use pyPDF2 to extract text from the resume
        parsed_user_data = generate_sections(resume_text) # use OpenAI to parse the resume text
        print(F"parsed_user_data: {parsed_user_data}")

        import sys
        from pathlib import Path
        # Add the sandbox directory to sys.path
        sys.path.append(str(Path(__file__).resolve().parent.parent / 'sandbox'))

        import mock_resumeAI as rad

        mock_event = {
            "body": json.dumps(parsed_user_data)
        }

        rad.convert_pdf_to_docx2(mock_event, None)
        return None
        
    else:
        filename = event_body.get('filename')
        resume_text = get_resume_text(filename) # use pyPDF2 to extract text from the resume
        parsed_user_data = generate_sections(resume_text) # use OpenAI to parse the resume text
        print(F"parsed_user_data: {parsed_user_data}")
    
    print(f"execution time took: {time.time() - execution_start} seconds.")
    
    response = {
            "statusCode": 200,
            "body": json.dumps(parsed_user_data)
        }
    return response

