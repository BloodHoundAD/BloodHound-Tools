# BloodHound Database Creator

This python script will generate a randomized data set for testing BloodHound features and analysis.

## Requirements

This script requires Python 3.7+, as well as the neo4j module. The script will only work with BloodHound 3.0.0 and above.

The Neo4j module can be installed using pip:

```
pip install neo4j
```

or

```
pip install -r requirements.txt
```

## Running

Ensure that all files in this repo are in the same directory.

```
python DBCreator.py
```

## Commands

- dbconfig - Set the credentials and URL for the database you're connecting too
- connect - Connects to the database using supplied credentials
- setnodes - Set the number of nodes to generate (defaults to 500, this is a safe number!)
- setdomain - Set the domain name
- cleardb - Clears the database and sets the schema properly
- generate - Generates random data in the database
- clear_and_generate - Connects to the database, clears the DB, sets the schema, and generates random data
- exit - Exits the script


## How to run DBCreator
Open a terminal and run the following commands:
```sh
$ git clone https://github.com/nicolas-carolo/BloodHound-Tools
$ cd BloodHound-Tools/DBCreator
$ python3 -m venv ./venv
$ source venv/bin/activate
$ pip3 install -r requirements.txt
$ python3 DBCreator.py
```
