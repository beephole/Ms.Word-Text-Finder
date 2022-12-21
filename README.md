# Ms.Word-Text-Finder



*Ms.Word-Text-Finder is a Python program that allows you to search for and extract specific text from a Microsoft Word document (.docx).*

Simply open the file and enter a search term to search for within the document. If the term is found, the program will print the 
paragraph containing the term and allow you to specify points within the paragraph to extract a portion of the text.
The extracted text can be saved as a variable and written to a text file. It can replace variables in a Word Document
or even converted the Word Template to a PDF file.




## *HOW TO INSTALL*

```      
git clone https://github.com/beephole/Ms.Word-Text-Finder.git
```
```
cd  Ms.Word-Text-Finder
```
```
pip install -r requirements.txt
```





##*Usage*


```
Ms.WordTxtFinder.py [-h] [-f FILE] [-d DIRECTORY] [-t TEMPLATE] [-o OUTPUT] [-H] [-b] [-pdf]
```




##**Examples**



To extract text from the file 'wordDocument.docx' and save it to a text file called 'output.txt':

```
python Ms.WordTxtFinder.py -f wordDocument.docx -o output.txt
```
To extract text from the file 'wordDocument.docx' and save it to'output.txt' and replace Variables at WordTemplate:

```
python Ms.WordTxtFinder.py -f wordDocument.docx -t template -o output.txt
```
To open a OS Window to get the file PAT ,search for text and save it to a text file called 'output.txt':

```
python Ms.WordTxtFinder.py -b -o output.txt
```
To extract text from the file 'wordDocument.docx' replace the Variables and also convertin to PDF':

```
python Ms.WordTxtFinder.py -f wordDocument.docx -t wordTemplate -pdf
```




    
![Command Prompt-tool](https://user-images.githubusercontent.com/118709832/208887256-8098754d-dc99-4c2e-a550-cb38de2d18d4.png)




##**Features**



   >1. Search for specific terms within a Microsoft Word document
   
   >2. Extract a portion of text from a paragraph
   
   >3. Save extracted text to a text file or Replace text with Word Template
   
   >4. Option to browse for a file using a Tk window
   
   >5. Customize the output file 

   >6. Converts a Word Template into a PDF




##**License**


Copyright (c) 2022 Beephole. This software is licensed under the MIT License. See the LICENSE file for details.




##**Contributing**


We welcome contributions to Text Extractor! If you have an idea for a new feature or have found a bug, 
please open an issue or submit a pull request.

 
   
   
   


*PAY ATTENTION*


1. Everything is saved in the working or current Directory.
2. If -d flag is not included then the PATH of files is going to be Desktop.
3. MAKE SURE WHEN USING FLAGS TO TYPE THE FILE NAMES CORRECTLY .




> "I’m always doing things I can’t do. That’s how I get to do them. :+1:"

> btc: 137L6AWxzsJ5eqsptGZx2yEfuznR9qntk3
