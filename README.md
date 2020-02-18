# Fill in templated document (docx) with an option to convert to PDF 

If you're like me, trying to apply to hundreds of companies to have them ask you to submit a cover letter. Then you have to change refill your templated word document with "recruiter X", "company Y", and "position Z". So much copy and paste hassle!

Look no further. I've put together this fill-in template document, with the optional conversion from document to pdf.



## How to use:
   
1. Have your template doc be in "docx" extension. The dependency python.docx does not work with "doc" extension. Under the hood, "docx" is a zip file containing xml and css file.

    **How to change "doc" to "docx":**

        Open the respective doc file in Microsoft Word.
        File -> Save As -> Select Word Document (*.docx) in dropdown menu -> Save
    
2. Fill your docx file with variables in form of ${name} format.
        
    **Example:**

        Dear ${recipient} of ${company}

3. Open profile.json and put in the variable names and values.

    >Note: profile.json is an array at the top level. You can put in as many profile as you want. The output(s) will be the name of the file + docx/pdf (apple.pdf, google.docx)

    >Note 2: if no matching ${key} is given, the program can optionally convert the docx to pdf file. View edit options below. The number of converted file corresponds to the number of profiles in profile.json

    **Example:**

        [
            {
                "output_name": "Apple",
                "profile": {
                    "my_name":"STEVE JOBLESS",
                    "MY_NAME": "STEVE JOBLESS",
                    "recipient": "Zukerman",
                    "company": "Tisla"
                }
            }
        ]

4. In settings.json. Give target_template the path to your template file. Including the name + extensions of the file.
   
   **Example:**

        Relative path: "/template.docx"
        Absolute path: "C:/Users/Tree/Desktop/template.docx"

5. For simple conversion to pdf. Keep "create_pdf" as true. To simply fill in the template without converting to pdf, switch to false. To keep a "docx" version of the filled template. Turn "create_word_doc" to true. 

    **Examples:**

        {
            "create_pdf": true,
            "create_word_doc": true,
            "target_template": "./sample_format/template.docx",
            "output_location": "./sample_output/"
        }
    >This will fill in the template, keep the filled-in version in .docx and make a pdf copy.

6. Run word_to_pdf.py

    >You will not lose template file with this program. Everything is a copy.

## Dependencies:

   - docx (to manipulate word document)
   - comptypes (to convert to pdf)
   - lxml (docx's dependency for handling the xml tree)

### For easy dependency installation:

    pip install -r requirements.txt

## How does this work?

This program works in 2 stages. First it fill in template with the key-pair values by running linear search and replace operations through the entire document. Save the filled in as a new docx document, then optionally convert said new version to pdf. 
