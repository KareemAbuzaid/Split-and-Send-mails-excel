# Split Excel sheet by category

## Split file

The first file `split_script` splits an excel sheet based on the different values in the column number in line **24** (column numbers start at 1)

## Send files

The second file `send_script` needs to run in a file that has the below format. Do not add 'xlsx' to the file names as the function already adds it.
By default the code creates the e-mail obejects then displays them. That can be changed by commenting the displaying line and uncommenting the sending line.
There is also an option to send the e-mail on behalf of another e-mail inbox that the current Outlook application has access to, by changing line number **84**


*All the files that both files access need to be in the same directory.*


File name | recipients |
--- | --- | 
filename_1 | test_1@example.com;test_2@example.com; |
filename_2 | test_3@example.com;test_4@example.com;test_5@example.com; |