# EXPERIMENT 4: DATA WRANGLING AND DATA VISUALIZATION

**Made by: Adrian Lorenz I. Dayao | 2ECE-A**

The objective of the experiment will be to fully utilize the pandas library and its connections to matplotlib for data processesing and interpretation. A sample file 'board2.xlsx' will be used as the source of the data that will be used in the different programs. Before the dataset is used, it is first prepared for processing by being read and saved as variable `board2`.

```python
import pandas as pd
board2 = pd.read_excel('board2.xlsx') #reads the xlsx excel file and saves as a variable with the name 'board2'
```
This will yield an output of:
|    | Name   | Gender   | Track            | Hometown   |   Math |   Electronics |   GEAS |   Communication |
|---:|:-------|:---------|:-----------------|:-----------|-------:|--------------:|-------:|----------------:|
|  0 | S1     | Male     | Instrumentation  | Luzon      |     58 |            89 |     75 |              78 |
|  1 | S2     | Female   | Communication    | Mindanao   |     52 |            75 |     90 |              52 |
|  2 | S3     | Female   | Instrumentation  | Mindanao   |     83 |            74 |     77 |              57 |
|  3 | S4     | Male     | Instrumentation  | Visayas    |     65 |            58 |     91 |              68 |
|  4 | S5     | Male     | Communication    | Luzon      |     59 |            86 |     43 |              88 |
|  5 | S6     | Female   | Microelectronics | Visayas    |     88 |            45 |     86 |              83 |
|  6 | S7     | Female   | Instrumentation  | Luzon      |     66 |            60 |     60 |              48 |
|  7 | S8     | Male     | Instrumentation  | Luzon      |     49 |            81 |     64 |              53 |
|  8 | S9     | Male     | Instrumentation  | Luzon      |     50 |            36 |     63 |              42 |
|  9 | S10    | Male     | Microelectronics | Mindanao   |     80 |            84 |     61 |              44 |
| 10 | S11    | Female   | Communication    | Visayas    |     48 |            56 |     48 |              67 |
| 11 | S12    | Male     | Communication    | Visayas    |     89 |            67 |     84 |              64 |
| 12 | S13    | Female   | Microelectronics | Luzon      |     88 |            35 |     83 |              43 |
| 13 | S14    | Female   | Microelectronics | Luzon      |     83 |            77 |     89 |              73 |
| 14 | S15    | Female   | Microelectronics | Mindanao   |     69 |            41 |     40 |              86 |
| 15 | S16    | Female   | Communication    | Luzon      |     71 |            70 |     87 |              81 |
| 16 | S17    | Female   | Microelectronics | Mindanao   |     81 |            79 |     77 |              45 |
| 17 | S18    | Male     | Communication    | Visayas    |     81 |            40 |     81 |              52 |
| 18 | S19    | Male     | Microelectronics | Luzon      |     79 |            63 |     79 |              71 |
| 19 | S20    | Female   | Communication    | Mindanao   |     59 |            60 |     62 |              85 |
| 20 | S21    | Female   | Microelectronics | Visayas    |     83 |            51 |     68 |              72 |
| 21 | S22    | Female   | Communication    | Visayas    |     64 |            39 |     89 |              58 |
| 22 | S23    | Male     | Instrumentation  | Luzon      |     84 |            70 |     74 |              47 |
| 23 | S24    | Female   | Microelectronics | Visayas    |     85 |            45 |     60 |              41 |
| 24 | S25    | Male     | Communication    | Luzon      |     74 |            91 |     94 |              42 |
| 25 | S26    | Female   | Instrumentation  | Visayas    |     71 |            47 |     83 |              62 |
| 26 | S27    | Male     | Microelectronics | Visayas    |     70 |            47 |     40 |              86 |
| 27 | S28    | Male     | Communication    | Visayas    |     85 |            53 |     80 |              53 |
| 28 | S29    | Male     | Instrumentation  | Mindanao   |     73 |            48 |     71 |              62 |
| 29 | S30    | Male     | Instrumentation  | Luzon      |     78 |            81 |     57 |              56 |

Afterwards a new column is appended wherein which it will contain each student's calculated averages from all subjects

```python
averages = board2[['Math', 'Electronics', 'GEAS', 'Communication']].mean(axis=1) #creates a new dataframe containing the averages of the chosen columns inside the bracket 
boardavg = pd.concat([board2, averages.rename('Average')], axis=1) #combines newly created dataframe of averages with board 2 to create an all in one dataframe to hold values.
boardavg
```

This will yield an output of:

|    | Name   | Gender   | Track            | Hometown   |   Math |   Electronics |   GEAS |   Communication |   Average |
|---:|:-------|:---------|:-----------------|:-----------|-------:|--------------:|-------:|----------------:|----------:|
|  0 | S1     | Male     | Instrumentation  | Luzon      |     58 |            89 |     75 |              78 |     75    |
|  1 | S2     | Female   | Communication    | Mindanao   |     52 |            75 |     90 |              52 |     67.25 |
|  2 | S3     | Female   | Instrumentation  | Mindanao   |     83 |            74 |     77 |              57 |     72.75 |
|  3 | S4     | Male     | Instrumentation  | Visayas    |     65 |            58 |     91 |              68 |     70.5  |
|  4 | S5     | Male     | Communication    | Luzon      |     59 |            86 |     43 |              88 |     69    |
|  5 | S6     | Female   | Microelectronics | Visayas    |     88 |            45 |     86 |              83 |     75.5  |
|  6 | S7     | Female   | Instrumentation  | Luzon      |     66 |            60 |     60 |              48 |     58.5  |
|  7 | S8     | Male     | Instrumentation  | Luzon      |     49 |            81 |     64 |              53 |     61.75 |
|  8 | S9     | Male     | Instrumentation  | Luzon      |     50 |            36 |     63 |              42 |     47.75 |
|  9 | S10    | Male     | Microelectronics | Mindanao   |     80 |            84 |     61 |              44 |     67.25 |
| 10 | S11    | Female   | Communication    | Visayas    |     48 |            56 |     48 |              67 |     54.75 |
| 11 | S12    | Male     | Communication    | Visayas    |     89 |            67 |     84 |              64 |     76    |
| 12 | S13    | Female   | Microelectronics | Luzon      |     88 |            35 |     83 |              43 |     62.25 |
| 13 | S14    | Female   | Microelectronics | Luzon      |     83 |            77 |     89 |              73 |     80.5  |
| 14 | S15    | Female   | Microelectronics | Mindanao   |     69 |            41 |     40 |              86 |     59    |
| 15 | S16    | Female   | Communication    | Luzon      |     71 |            70 |     87 |              81 |     77.25 |
| 16 | S17    | Female   | Microelectronics | Mindanao   |     81 |            79 |     77 |              45 |     70.5  |
| 17 | S18    | Male     | Communication    | Visayas    |     81 |            40 |     81 |              52 |     63.5  |
| 18 | S19    | Male     | Microelectronics | Luzon      |     79 |            63 |     79 |              71 |     73    |
| 19 | S20    | Female   | Communication    | Mindanao   |     59 |            60 |     62 |              85 |     66.5  |
| 20 | S21    | Female   | Microelectronics | Visayas    |     83 |            51 |     68 |              72 |     68.5  |
| 21 | S22    | Female   | Communication    | Visayas    |     64 |            39 |     89 |              58 |     62.5  |
| 22 | S23    | Male     | Instrumentation  | Luzon      |     84 |            70 |     74 |              47 |     68.75 |
| 23 | S24    | Female   | Microelectronics | Visayas    |     85 |            45 |     60 |              41 |     57.75 |
| 24 | S25    | Male     | Communication    | Luzon      |     74 |            91 |     94 |              42 |     75.25 |
| 25 | S26    | Female   | Instrumentation  | Visayas    |     71 |            47 |     83 |              62 |     65.75 |
| 26 | S27    | Male     | Microelectronics | Visayas    |     70 |            47 |     40 |              86 |     60.75 |
| 27 | S28    | Male     | Communication    | Visayas    |     85 |            53 |     80 |              53 |     67.75 |
| 28 | S29    | Male     | Instrumentation  | Mindanao   |     73 |            48 |     71 |              62 |     63.5  |
| 29 | S30    | Male     | Instrumentation  | Luzon      |     78 |            81 |     57 |              56 |     68    |

And serve as the basis for data to be used in the upcoming programs.

# **A. VISAYAS COMMUNICATION DATAFRAME**

a.) For the first program, the dataset is filtered to only save rows whose Hometown element is `Visayas` and whose Track element is `Communications`. 

```python
board3 =boardavg[(boardavg['Hometown'] == 'Visayas') #filters to only show rows categorized into visayas for the hometown column
     & 
     (boardavg['Track'] == 'Communication')] #specifies further that the chosen rows must also be categorized as communication in the track column
print(board3.to_markdown())
```

Test output with print command:
```python
|    | Name   | Gender   | Track         | Hometown   |   Math |   Electronics |   GEAS |   Communication |   Average |
|---:|:-------|:---------|:--------------|:-----------|-------:|--------------:|-------:|----------------:|----------:|
| 10 | S11    | Female   | Communication | Visayas    |     48 |            56 |     48 |              67 |     54.75 |
| 11 | S12    | Male     | Communication | Visayas    |     89 |            67 |     84 |              64 |     76    |
| 17 | S18    | Male     | Communication | Visayas    |     81 |            40 |     81 |              52 |     63.5  |
| 21 | S22    | Female   | Communication | Visayas    |     64 |            39 |     89 |              58 |     62.5  |
| 27 | S28    | Male     | Communication | Visayas    |     85 |            53 |     80 |              53 |     67.75 |
```

Afterwards only the columns Name, Gender, Math, Electronics, and Average are retained for the new dataset variable `VisComm` :

```python
VisComm = board3.loc[:, ['Name', 'Gender', 'Math', 'Electronics', 'Average']] #only saves the chosen columns into 'VisComm' variable
print(VisComm.to_markdown()) 
```

Test output with print command:
```python
|    | Name   | Gender   |   Math |   Electronics |   Average |
|---:|:-------|:---------|-------:|--------------:|----------:|
| 10 | S11    | Female   |     48 |            56 |     54.75 |
| 11 | S12    | Male     |     89 |            67 |     76    |
| 17 | S18    | Male     |     81 |            40 |     63.5  |
| 21 | S22    | Female   |     64 |            39 |     62.5  |
| 27 | S28    | Male     |     85 |            53 |     67.75 |
```

```python
print('Amount of rows:', VisComm.shape[0]) #show amount of rows in VisComm
```

Output:
`Amount of rows: 5`

# **B. Visayas Female Dataframe**

a.) For the first program, the dataset is filtered to only save rows whose Hometown element is `Visayas` and whose Gender element is `Female`, of which afterwards only the columns Name, Gender, Math, Electronics, and Average are retained for the new dataset variable. 

```python
VisFemale =boardavg[(boardavg['Hometown'] == 'Visayas') #filters to only show rows categorized into visayas for the hometown column
     & 
     (boardavg['Gender'] == 'Female')].loc[:, ['Name', 'Track', 'GEAS', 'Electronics', 'Average']] #specifies filter to also only show rows categorized into female for the Gender column while choosing to only display chosen columns
print(VisFemale.to_markdown()) 
```

Test output with print command:
```python
|    | Name   | Track            |   GEAS |   Electronics |   Average |
|---:|:-------|:-----------------|-------:|--------------:|----------:|
|  5 | S6     | Microelectronics |     86 |            45 |     75.5  |
| 10 | S11    | Communication    |     48 |            56 |     54.75 |
| 20 | S21    | Microelectronics |     68 |            51 |     68.5  |
| 21 | S22    | Communication    |     89 |            39 |     62.5  |
| 23 | S24    | Microelectronics |     60 |            45 |     57.75 |
| 25 | S26    | Instrumentation  |     83 |            47 |     65.75 |
```

`VisFemale` is then treated to only show rows whose average is atleast at the value of 60:

```python
VisFemale[VisFemale['Average']>=60] #creates dataframe that excludes any element with Average lower than 60 from VisFemale dataframe
```

Test output:
```python
|    | Name   | Track            |   GEAS |   Electronics |   Average |
|---:|:-------|:-----------------|-------:|--------------:|----------:|
|  5 | S6     | Microelectronics |     86 |            45 |     75.5  |
| 20 | S21    | Microelectronics |     68 |            51 |     68.5  |
| 21 | S22    | Communication    |     89 |            39 |     62.5  |
| 25 | S26    | Instrumentation  |     83 |            47 |     65.75 |
```

# **C. Category-Average Visualization**
Average comparison from Hometown, Gender, and Track columns

a.) The First set of code calculates the mean of the Averages per column to be observed.

```python
boardavg.groupby('Track')['Average'].mean() #Gets the mean value for the Average column of each category in Track
```
Test output:
Track
Communication       67.975
Instrumentation     65.225
Microelectronics    67.500
Name: Average, dtype: float64

```python
boardavg.groupby('Gender')['Average'].mean() #Gets the mean value for the Average column of each category in Gender
```
Test output:
Gender
Female    66.616667
Male      67.183333
Name: Average, dtype: float64
Selection deleted

```python
boardavg.groupby('Hometown')['Average'].mean() #Gets the mean value for the Average column of each category in Hometown
```

Test output:
Hometown
Luzon       68.083333
Mindanao    66.678571
Visayas     65.750000
Name: Average, dtype: float64

b.) This part of the code then saves and displays these calculated as dataframe variables.

```python
Trackavg = pd.DataFrame(data=boardavg.groupby('Track')['Average'].mean()) #turns the calculated means of Track category from boardavg and creates then displays its dataframe while saving to a variable
```

Test output:
| Track            |   Average |
|:-----------------|----------:|
| Communication    |    67.975 |
| Instrumentation  |    65.225 |
| Microelectronics |    67.5   |

```python
Genderavg = pd.DataFrame(data=boardavg.groupby('Gender')['Average'].mean()) #turns the calculated means of Gender category from boardavg and creates then displays its dataframe while saving to a variable
```

Test output:
| Gender   |   Average |
|:---------|----------:|
| Female   |   66.6167 |
| Male     |   67.1833 |

```python
Hometownavg = pd.DataFrame(data=boardavg.groupby('Hometown')['Average'].mean()) #turns the calculated means of Hometown category from boardavg and creates then displays its dataframe while saving to a variable
```

Test output:
| Hometown   |   Average |
|:-----------|----------:|
| Luzon      |   68.0833 |
| Mindanao   |   66.6786 |
| Visayas    |   65.75   |

c.) The second part of the program creates one figure containing three bar charts: mean Average by Track, by Gender, and by Hometown.
The calculated dataframes from a. are combined into one dataframe whose columns are reorganized for a cleaner graph display
```python
categories = pd.concat([Hometownavg, Genderavg, Trackavg], axis = 1) #combines all mean values from the saved dataframe variables into one new dataframe
categories.columns = ['Hometown averages', 'Gender averages', 'Track averages'] #assigns columns to each group of averages to sort by category
```
The graph is then designed with dimensions, type, title, and values being specified.
```python
category_average = categories.plot(kind='bar', layout=(1, 3), figsize=(12, 4), ylim = (60, 70), rot=0, title="Mean Averages per category") 
#creates a figure of bar graphs from the categories dataframe with a minimum vertical value of 60 and max value of 70
```

Test output:
<img width="4800" height="1600" alt="category average" src="https://github.com/user-attachments/assets/b666be45-4aa6-42e5-b0ee-e1a53c0fc3e2" />


d.) With the graph completed, the given dataset can then be interpreted.

## Data interpretations:
Hometown Averages: For the feature of hometown averages, the graph shows that Luzon has the highest mean among the three categories.
Gender Averages: Between the two categories in the gender averages feature, the male category has a higher mean value for its averages.
Track Averages: The communications category contains the highest mean value for its averages in the Track feature as compared to the other two categories.
