# xls2object
This little project aims to load data from excel and populate objects from an external assembly and bring those objects to the client of this toolkit.

Once i was asked to implement excel file imports to a system. So, firstable, i've asked for the source spreadsheet and its respective database model destination. To my surprise, none o them was already defined. So i had to think in a way to implement this functionality being able to read any excel file and transform its data into any kind of object i could once this be properly defined, and even able to adjust this process with fewer changes as possible in order to make it maintainable at lower possible cost.

The solution i've decide to apply was:
- Create a json file with the rules to understand and read any excel file
- In the same json file map the attributes of an object of an unknown and not referenced(by this toolkit) assembly to each excel cell.
- Use reflection to create an instance of this class and populate its attributes with data found in the excel file according to the rules defined in the json file. 
