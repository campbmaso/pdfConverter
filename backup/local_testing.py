if event_body.get("local_testing"):
        resume_text = ""
        with open(
            "/Users/campbmaso/Desktop/Development/GitHub/pdfConverter/resumes/Cory Mazure - 2023 Professional Resume.pdf",
            "rb",
        ) as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                resume_text += page.extract_text()
        import sys
        from pathlib import Path
        # Add the sandbox directory to sys.path
        sys.path.append(str(Path(__file__).resolve().parent.parent / 'sandbox'))

        import resumeAI_dummy as rad

        basic_event = {
            "key1": "value1",
            "key2": "value2",
            "key3": "value3"
        }

        rad.convert_pdf_to_docx2(basic_event, None)
        return None