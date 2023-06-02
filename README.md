# JB4-Log-Calculator
A program to assist with getting useful data from JB4 log files.

## Example:
### Data
Data for each Map -> Gear -> RPM ( floored to nearest 500). All data is averaged across all log files per map. 
 ![image](https://github.com/jpcoker3/JB4-Log-Calculator/assets/111995337/52625517-ca75-4d76-be36-ad0ee7b2598b)

### Graphs
Graphs for each map. Fun to look at. Not super useful in its current form but I'm looking to change it in the future.

Data shown is for each data point across each map file, not averaged or cleaned or anything. 
![image](https://github.com/jpcoker3/JB4-Log-Calculator/assets/111995337/982e3924-8903-45ef-9556-ad6d3268e8e5)


## Setup
( this assumes you have Python installed. a .exe version of this project may come along in the future if there is any demand at all) 

For this program to work, you must install pandas and openpyxl. 

Enter the following in your terminal. this assumes you have [pip](https://pip.pypa.io/en/stable/installation/) installed. If you are struggling, look up how to install these on your particular setup. 

```
pip install pandas
pip install openpyxl
```

Now, before running you must create 2 things. a folder containing all of your logs (in either .xlsx or .csv), and an excel file. 

![image](https://github.com/jpcoker3/JB4-Log-Calculator/assets/111995337/f1cb55ac-d963-45ab-836f-c56d7724d2b7)


Lastly, in the main() function, replace the strings with the path to the file and folder as shown below. 

![image](https://github.com/jpcoker3/JB4-Log-Calculator/assets/111995337/b792c6d5-e3ad-41a4-b73d-5d29efc96d91)

Note that on Windows, you can click a file and then ctrl+shift+c to copy the path. the "r" infront of the path is needed if you do it this way. 


## Complete!
And thats it! you should now be able to run the program. Note that the files must be closed for the program to run.

## Other

Feel free to message me about any suggestions or questions with this program.
