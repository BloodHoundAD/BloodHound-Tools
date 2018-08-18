# BloodHound Database Creator
This python script will generate a randomized data set for testing BloodHound features and analysis. 

## Requirements
This script requires Python 2.7, as well as the neo4j-driver. The script will only work with BloodHound 2.0.3 and above.

The Neo4j Driver can be installed using pip:

```
pip install neo4j-driver
```

## Running
Ensure that all files in this repo are in the same directory.

```
python DBCreator.py
```

## Commands
* dbconfig - Set the credentials and URL for the database you're connecting too
* connect - Connects to the database using supplied credentials
* setnodes - Set the number of nodes to generate (defaults to 500, this is a safe number!)
* cleardb - Clears the database and sets the schema properly
* generate - Generates random data in the database
* clear_and_generate - Connects to the database, clears the DB, sets the schema, and generates random data
* exit - Exits the script
