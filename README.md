# Parser Operation Logic

- Place the parser and the table file in the same folder.
- Upon the first run, the program will create a settings.json file, specifying default parameters for column numbers.
- When generating a report, a Reports folder is created in the current directory to store report files.
- The naming convention for reports is as follows: {company name}_{company address}.docx
- If a report file with the same name as the one being generated already exists in the Reports folder, the old file is moved to the trash, and a new one is created.
