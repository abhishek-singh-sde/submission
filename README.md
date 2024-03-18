Create a virtualenv and install the libraries from requirements.txt

Now, to specify path to the input file:
py final.py --path=<Path_To_File>

Eg.
>> py final.py --path="Test PDF.pdf"

After running this command,
A. The extraction of input file shall take place
B. The cleaning shall be performed on the extracted text
C. The extracted/generated text shall be stored in folder InputFileHistory for future ref 
D. The file Final_Updated_Records.xlsx shall be updated, making necessary additions and taking care of deduplications

Some sample calls to the reporting functions have also been made which will show up once the script runs.

NOTE: Post the first run, we can notice the values getting added in the (initially empty) file named "Final_Updated_Records.xlsx". In the subsequent runs, if try to upload the duplicate copy, i.e. rerun final.py --path=<TestPDF's path> then we will observe that no additional entries will be made and Final_Updated_Records correctly remains as it is.
