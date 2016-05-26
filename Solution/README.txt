The main program is 'pdf2text.py' which makes use of cStringIO and pdfMiner libraries to extract characters from the pdf 'test.pdf' and print it to 'output.txt'.

Then 'text2list.py' makes use of Regular expressions and parses data from 'output.txt' to "Formatted_Parsed_Output.xlsx".

There are many a peculiaritis in the 'test.pdf' file which can be handled by 'text2list.py' during parsing. Hence for any bank statement of HDFC Bank, this solution will work without any modifications.

Also, for other banks such as ICICI, AXIS BANK, etc. only parsing will have to be changed. Main file 'pdf2text.py' will not require any modifications.