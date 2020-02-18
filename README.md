# Fill in templated document (docx)/ doc to PDF converter

If you're like me, trying to apply to hundreds of companies to have them ask you to submit a cover letter. Then you have to change refill your templated word document with "recruiter X", "company Y", and "position Z". So much copy and paste hassle.

Look no further. I've put together this fill-in template document, with the optional conversion from document to pdf.



## How to use:

1. First, have your template doc be in "docx" extension. The dependency python.docx does not work with doc extension. Under the hood, docx is a zip file containing xml and css file.

    **How to change to docx:**

        Open the respective doc file in Microsoft Word.
        File -> Save As -> Select docx in dropdown menu -> Save
    
2. Fill your template with variables in ${variable_name} format.
        
    **Example:**

        Dear ${recipient} of ${company}

3. Open profile.json and put in the variable names and values

    >Note: if no matching ${key} is given, the program can optionally convert the docx to pdf file

    **Example:**

        {
            "recipient": "Steve soon-to-get-Jobs",
            "company" : "Apple"
        }

4. In settings.json. Give template_location the path to your template file. Including the name + extensions of the file.
   
   **Example:**

        Relative path: "/template.docx"
        Absolute path: "C:/Users/Tree/Desktop/template.docx"

5. For simple conversion to pdf. Keep "convert_to_pdf": true. To simply fill in the template, switch it to false. To keep a .docx copy of the filled version. Turn "keep_word_copy" to true. 

    **Examples:**

        To fill in my template, and convert it to pdf

6. Run word_to_pdf_bot.py


>Note: this program works in 2 stages. First it fill in template with the key-pair values by running linear search and replace operations through the entire document. Save the filled in as a new docx document, then optionally convert said new version to pdf. 

## Dependencies:

   - docx (to manipulate word document)
   - comptypes (to convert to pdf)
   - lxml (docx's dependency for handling the xml tree)

**For easy dependency installation:**

    pip install -r requirements.txt
