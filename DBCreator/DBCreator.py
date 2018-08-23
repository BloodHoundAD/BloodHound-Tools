# Requirements - pip install neo4j-driver
# This script is used to create randomized sample databases. 
# Commands
# 	dbconfig - Set the credentials and URL for the database you're connecting too
#	connect - Connects to the database using supplied credentials
# 	setnodes - Set the number of nodes to generate (defaults to 500, this is a safe number!)
# 	cleardb - Clears the database and sets the schema properly
#	generate - Generates random data in the database
#	clear_and_generate - Connects to the database, clears the DB, sets the schema, and generates random data

from neo4j.v1 import GraphDatabase
import cmd
import os
import sys
import random
import cPickle
import math
import itertools
from collections import defaultdict
import uuid
import time

class Messages():
	def title(self):
		print "================================================================"
		print "BloodHound Sample Database Creator"
		print "================================================================"

	def input_default(self, prompt, default):
		return raw_input("%s [%s] " % (prompt, default)) or default

class MainMenu(cmd.Cmd):
	def __init__(self):
		self.m = Messages()
		self.url = "bolt://localhost:7687"
		self.username = "neo4j"
		self.password = "neo4jj"
		self.driver = None
		self.connected = False
		self.num_nodes = 500
		self.current_time = int(time.time())
		self.base_sid = "S-1-5-21-883232822-274137685-4173207997"
		with open('first.pkl', 'rb') as f:
			self.first_names = cPickle.load(f)
		
		with open('last.pkl', 'rb') as f:
			self.last_names = cPickle.load(f)
		
		cmd.Cmd.__init__(self)
	
	def cmdloop(self):
		while True:
			self.m.title()
			self.do_help("")
			try:
				cmd.Cmd.cmdloop(self)
			except KeyboardInterrupt:
				if self.driver is not None:
					self.driver.close()
				raise KeyboardInterrupt

	def help_dbconfig(self):
		print "Configure database connection parameters"

	def help_connect(self):
		print "Test connection to the database and verify credentials"

	def help_setnodes(self):
		print "Set base number of nodes to generate (default 500)"

	def help_cleardb(self):
		print "Clear the database and set constraints"

	def help_generate(self):
		print "Generate random data"

	def help_clear_and_generate(self):
		print "Connect to the database, clear the db, set the schema, and generate random data"

	def help_exit(self):
		print "Exits the database creator"

	def do_dbconfig(self, args):
		print "Current Settings:"
		print "DB Url: {}".format(self.url)
		print "DB Username: {}".format(self.username)
		print "DB Password: {}".format(self.password)
		print ""
		self.url = self.m.input_default("Enter DB URL", self.url)
		self.username = self.m.input_default("Enter DB Username", self.username)
		self.password = self.m.input_default("Enter DB Password", self.password)
		print ""
		print "New Settings:"
		print "DB Url: {}".format(self.url)
		print "DB Username: {}".format(self.username)
		print "DB Password: {}".format(self.password)
		print ""
		print "Testing DB Connection"
		self.test_db_conn()

	def do_setnodes(self, args):
		passed = args
		if passed != "":
			try:
				self.num_nodes = int(passed)
				return
			except ValueError:
				pass

		self.num_nodes = int(self.m.input_default("Number of nodes of each type to generate", self.num_nodes))

	def do_exit(self, args):
		raise KeyboardInterrupt

	def do_connect(self, args):
		self.test_db_conn()

	def do_cleardb(self, args):
		if not self.connected:
			print "Not connected to database. Use connect first"
			return
		
		print "Clearing Database"
		d = self.driver
		session = d.session()
		num = 1
		while num > 0:
			result = session.run("MATCH (n) WITH n LIMIT 10000 DETACH DELETE n RETURN count(n)")
			for r in result:
				num = int(r['count(n)'])

		print "Resetting Schema"
		for constraint in session.run("CALL db.constraints"):
			session.run("DROP {}".format(constraint['description']))

		for index in session.run("CALL db.indexes"):
			session.run("DROP {}".format(index['description']))

		session.run("CREATE CONSTRAINT ON (c:User) ASSERT c.name IS UNIQUE")
		session.run("CREATE CONSTRAINT ON (c:Group) ASSERT c.name IS UNIQUE")
		session.run("CREATE CONSTRAINT ON (c:Computer) ASSERT c.name IS UNIQUE")
		session.run("CREATE CONSTRAINT ON (c:Domain) ASSERT c.name IS UNIQUE")
		session.run("CREATE CONSTRAINT ON (c:OU) ASSERT c.guid IS UNIQUE")
		session.run("CREATE CONSTRAINT ON (c:GPO) ASSERT c.name IS UNIQUE")

		print "DB Cleared and Schema Set"

	def test_db_conn(self):
		self.connected = False
		if self.driver is not None:
			self.driver.close()
		try:
			self.driver = GraphDatabase.driver(self.url, auth=(self.username,self.password))
			self.connected = True
			print "Database Connection Successful!"
		except:
			self.connected = False
			print "Database Connection Failed. Check your settings."
	
	def do_generate(self, args):
		self.generate_data()
		
	def do_clear_and_generate(self, args):
		self.test_db_conn()
		self.do_cleardb("a")
		self.generate_data()

	def split_seq(self, iterable, size):
		it = iter(iterable)
		item = list(itertools.islice(it, size))
		while item:
			yield item
			item = list(itertools.islice(it, size))
	
	def generate_timestamp(self):
		choice = random.randint(-1,1)
		if choice == 1:
			variation = random.randint(0,31536000)
			return self.current_time - variation
		else:
			return choice
		

	def generate_data(self):
		if not self.connected:
			print "Not connected to database. Use connect first"
			return
		
		computers = []
		groups = []
		users = []
		gpos = []
		ou_guid_map = {}
		
		used_states = []

		states = ["WA","MD","AL","IN","NV","VA","CA","NY","TX","FL"]
		partitions = ["IT","HR","MARKETING","OPERATIONS","BIDNESS"]
		os_list = ["Windows Server 2003"] * 1 + ["Windows Server 2008"] * 15 + ["Windows 7"] * 35 + ["Windows 10"] * 28 + ["Windows XP"] * 1 + ["Windows Server 2012"] * 8 + ["Windows Server 2008"] * 12
		session = self.driver.session()

		print "Starting data generation with nodes={}".format(self.num_nodes)

		print "Populating Standard Nodes"
		session.run("MERGE (n:Group {name:'DOMAIN ADMINS@TESTLAB.LOCAL'}) SET n.highvalue=true,n.objectsid={sid}", sid=self.base_sid + "-512")
		session.run("MERGE (n:Group {name:'DOMAIN COMPUTERS@TESTLAB.LOCAL'})")
		session.run("MERGE (n:Group {name:'DOMAIN USERS@TESTLAB.LOCAL'})")
		session.run("MERGE (n:Group {name:'DOMAIN CONTROLLERS@TESTLAB.LOCAL'}) SET n.highvalue=true")
		session.run("MERGE (n:Group {name:'ENTERPRISE DOMAIN CONTROLLERS@TESTLAB.LOCAL'}) SET n.highvalue=true")
		session.run("MERGE (n:Group {name:'ENTERPRISE READ-ONLY DOMAIN CONTROLLERS@TESTLAB.LOCAL'})")
		session.run("MERGE (n:Group {name:'ADMINISTRATORS@TESTLAB.LOCAL'}) SET n.highvalue=true")
		session.run("MERGE (n:Group {name:'ENTERPRISE ADMINS@TESTLAB.LOCAL'}) SET n.highvalue=true")
		session.run("MERGE (n:Domain {name:'TESTLAB.LOCAL'})")
		ddp = str(uuid.uuid4())
		ddcp = str(uuid.uuid4())
		dcou = str(uuid.uuid4())
		session.run("MERGE (n:GPO {name:'DEFAULT DOMAIN POLICY@TESTLAB.LOCAL', guid:{guid}})", guid=ddp)
		session.run("MERGE (n:GPO {name:'DEFAULT DOMAIN CONTROLLERS POLICY@TESTLAB.LOCAL', guid:{guid}})", guid=ddcp)
		session.run("MERGE (n:OU {name:'DOMAIN CONTROLLERS@TESTLAB.LOCAL', guid:{guid}, blocksInheritance: false})", guid=dcou)
		

		print "Adding Standard Edges"
		
		#Default GPOs
		session.run('MERGE (n:GPO {name:"DEFAULT DOMAIN POLICY@TESTLAB.LOCAL"}) MERGE (m:Domain {name:"TESTLAB.LOCAL"}) MERGE (n)-[:GpLink {isacl:false, enforced:toBoolean(false)}]->(m)')
		session.run('MERGE (n:Domain {name:"TESTLAB.LOCAL"}) MERGE (m:OU {guid:{guid}}) MERGE (n)-[:Contains {isacl:false}]->(m)', guid=dcou)
		session.run('MERGE (n:GPO {name:"DEFAULT DOMAIN CONTROLLERS POLICY@TESTLAB.LOCAL"}) MERGE (m:OU {guid:{guid}}) MERGE (n)-[:GpLink {isacl:false, enforced:toBoolean(false)}]->(m)', guid=dcou)

		#Ent Admins -> Domain Node
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) MERGE (m:Group {name:"ENTERPRISE ADMINS@TESTLAB.LOCAL"}) MERGE (m)-[:GenericAll {isacl:true}]->(n)')
		
		#Administrators -> Domain Node
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) MERGE (m:Group {name:"ADMINISTRATORS@TESTLAB.LOCAL"}) MERGE (m)-[:Owns {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ADMINISTRATORS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:WriteOwner {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ADMINISTRATORS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:WriteDacl {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ADMINISTRATORS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:DCSync {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ADMINISTRATORS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChanges {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ADMINISTRATORS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChangesAll {isacl:true}]->(n)')
		
		#Domain Admins -> Domain Node
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:WriteOwner {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:WriteDacl {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:DCSync {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChanges {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChangesAll {isacl:true}]->(n)')
		
		#DC Groups -> Domain Node
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ENTERPRISE DOMAIN CONTROLLERS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChanges {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"ENTERPRISE READ-ONLY DOMAIN CONTROLLERS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChanges {isacl:true}]->(n)')
		session.run(
			'MERGE (n:Domain {name:"TESTLAB.LOCAL"}) WITH n MERGE (m:Group {name:"DOMAIN CONTROLLERS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:GetChangesAll {isacl:true}]->(n)')

		print "Generating Computer Nodes"
		props = []
		for i in xrange(1,self.num_nodes+1):
			comp_name = "COMP{:05d}.TESTLAB.LOCAL".format(i)
			computers.append(comp_name)
			os = random.choice(os_list)
			enabled = True
			props.append({'name':comp_name, 'props':{
				'operatingsystem':os,
				'enabled':enabled
			}})

			if (len(props) > 500):
				session.run('UNWIND {props} as prop MERGE (n:Computer {name:prop.name}) SET n += prop.props WITH n MERGE (m:Group {name:"DOMAIN COMPUTERS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)
				props = []
		session.run('UNWIND {props} as prop MERGE (n:Computer {name:prop.name}) WITH n MERGE (m:Group {name:"DOMAIN COMPUTERS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)

		print "Creating Domain Controllers"
		for state in states:
			comp_name = "{}LABDC.TESTLAB.LOCAL".format(state)
			session.run(
				'MERGE (n:Computer {name:{name}}) WITH n MERGE (m:Group {name:"DOMAIN CONTROLLERS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', name=comp_name)
			session.run(
				'MERGE (n:Computer {name:{name}}) WITH n MATCH (m:OU {name:"DOMAIN CONTROLLERS@TESTLAB.LOCAL",guid:{dcou}}) WITH n,m MERGE (m)-[:Contains]->(n)', name=comp_name, dcou=dcou)
			session.run(
				'MERGE (n:Computer {name:{name}}) WITH n MERGE (m:Group {name:"ENTERPRISE DOMAIN CONTROLLERS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', name=comp_name)
			session.run(
				'MERGE (n:Computer {name:{name}}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:AdminTo]->(n)', name=comp_name)

		
		used_states = list(set(used_states))

		print "Generating User Nodes"
		props = []
		for i in xrange(1, self.num_nodes+1):
			first = random.choice(self.first_names)
			last = random.choice(self.last_names)
			user_name = "{}{}{:05d}@TESTLAB.LOCAL".format(first[0], last,i).upper()
			users.append(user_name)
			dispname = "{} {}".format(first,last)
			enabled = True
			pwdlastset = self.generate_timestamp()
			lastlogon = self.generate_timestamp()
			sidint = i + 1000
			objectsid = self.base_sid + "-" + str(sidint)
			
			props.append({'name':user_name, 'props': {
				'displayname': dispname,
				'enabled': enabled,
				'pwdlastset': pwdlastset,
				'lastlogon': lastlogon,
				'objectsid': objectsid
			}})
			if (len(props) > 500):
				session.run(
					'UNWIND {props} as prop MERGE (n:User {name:prop.name}) SET n += prop.props WITH n MERGE (m:Group {name:"DOMAIN USERS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)
				props = []
		session.run(
			'UNWIND {props} as prop MERGE (n:User {name:prop.name}) SET n += prop.props WITH n MERGE (m:Group {name:"DOMAIN USERS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)

		
		print "Generating Group Nodes"
		weighted_parts = ["IT"] * 7 + ["HR"] * 13 + ["MARKETING"] * 30 + ["OPERATIONS"] * 20 + ["BIDNESS"] * 30
		props = []
		for i in xrange(1, self.num_nodes + 1):
			dept = random.choice(weighted_parts)
			group_name = "{}{:05d}@TESTLAB.LOCAL".format(dept,i)
			groups.append(group_name)
			props.append({'name':group_name})
			if len(props) > 500:
				session.run('UNWIND {props} as prop MERGE (n:Group {name:prop.name})', props=props)
				props = []
		
		session.run(
			'UNWIND {props} as prop MERGE (n:Group {name:prop.name})', props=props)
		
		print "Adding Domain Admins to Local Admins of Computers"
		props = []
		for comp in computers:
			props.append({'name':comp})
			if len(props) > 500:
				session.run('UNWIND {props} as prop MERGE (n:Computer {name:prop.name}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:AdminTo]->(n)', props=props)
				props = []
		
		session.run(
			'UNWIND {props} as prop MERGE (n:Computer {name:prop.name}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:AdminTo]->(n)', props=props)
		
		dapctint = random.randint(3,5)
		dapct = float(dapctint) / 100
		danum = int(math.ceil(self.num_nodes * dapct))
		danum = min([danum, 30])
		print "Creating {} Domain Admins ({}% of users capped at 30)".format(danum, dapctint)
		das = random.sample(users, danum)
		for da in das:
			session.run(
				'MERGE (n:User {name:{name}}) WITH n MERGE (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH n,m MERGE (n)-[:MemberOf]->(m)', name=da)

		print "Applying random group nesting"
		max_nest = int(round(math.log10(self.num_nodes)))
		props = []
		for group in groups:
			if (random.randrange(0,100) < 10):
				num_nest = random.randrange(1, max_nest)
				dept = group[0:-19]
				dpt_groups = [x for x in groups if dept in x]
				if num_nest > len(dpt_groups):
					num_nest = random.randrange(1, len(dpt_groups))
				to_nest = random.sample(dpt_groups, num_nest)
				for g in to_nest:
					if not g == group:
						props.append({'a':group,'b':g})
				
			if (len(props) > 500):
				session.run('UNWIND {props} AS prop MERGE (n:Group {name:prop.a}) WITH n,prop MERGE (m:Group {name:prop.b}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)
				props = []
		
		session.run('UNWIND {props} AS prop MERGE (n:Group {name:prop.a}) WITH n,prop MERGE (m:Group {name:prop.b}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)
		
		print "Adding users to groups"
		props = []
		a = math.log10(self.num_nodes)
		a = math.pow(a,2)
		a = math.floor(a)
		a = int(a)
		num_groups_base = a
		variance = int(math.ceil(math.log10(self.num_nodes)))
		it_users = []

		print "Calculated {} groups per user with a variance of - {}".format(num_groups_base,variance*2)

		for user in users:
			dept = random.choice(weighted_parts)
			if dept == "IT":
				it_users.append(user)
			possible_groups = [x for x in groups if dept in x]
			
			sample = num_groups_base + random.randrange(-(variance*2), 0)
			if (sample > len(possible_groups)):
				sample = int(math.floor(float(len(possible_groups)) / 4))
			
			if (sample == 0):
				continue
			
			to_add = random.sample(possible_groups, sample)
			
			for group in to_add:
				props.append({'a':user,'b':group})
			
			if len(props) > 500:
				session.run('UNWIND {props} AS prop MERGE (n:User {name:prop.a}) WITH n,prop MERGE (m:Group {name:prop.b}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)
				props = []
		
		session.run(
			'UNWIND {props} AS prop MERGE (n:User {name:prop.a}) WITH n,prop MERGE (m:Group {name:prop.b}) WITH n,m MERGE (n)-[:MemberOf]->(m)', props=props)
		
		it_users = it_users + das
		it_users = list(set(it_users))

		print "Adding local admin rights"
		it_groups = [x for x in groups if "IT" in x]
		random.shuffle(it_groups)
		super_groups = random.sample(it_groups, 4)
		super_group_num = int(math.floor(len(computers) * .85))

		it_groups = [x for x in it_groups if not x in super_groups]

		total_it_groups = len(it_groups)

		dista = int(math.ceil(total_it_groups * .6))
		distb = int(math.ceil(total_it_groups * .3))
		distc = int(math.ceil(total_it_groups * .07))
		distd = int(math.ceil(total_it_groups * .03))

		distribution_list = [1] * dista + [2] * distb + [10] * distc + [50] * distd

		props = []
		for x in xrange(0, total_it_groups):
			g = it_groups[x]
			dist = distribution_list[x]

			to_add = random.sample(computers, dist)
			for a in to_add:
				props.append({'a': g, 'b': a})

			if len(props) > 500:
				session.run('UNWIND {props} AS prop MERGE (n:Group {name:prop.a}) WITH n,prop MERGE (m:Computer {name:prop.b}) WITH n,m MERGE (n)-[:AdminTo]->(m)', props=props)
				props = []
		
		for x in super_groups:
			for a in random.sample(computers, super_group_num):
				props.append({'a': x, 'b': a})
			
			if len(props) > 500:
				session.run('UNWIND {props} AS prop MERGE (n:Group {name:prop.a}) WITH n,prop MERGE (m:Computer {name:prop.b}) WITH n,m MERGE (n)-[:AdminTo]->(m)', props=props)
				props = []

		session.run('UNWIND {props} AS prop MERGE (n:Group {name:prop.a}) WITH n,prop MERGE (m:Computer {name:prop.b}) WITH n,m MERGE (n)-[:AdminTo]->(m)', props=props)
		
		print "Adding RDP/ExecuteDCOM/AllowedToDelegateTo"
		count = int(math.floor(len(computers) * .1))
		props = []
		for i in xrange(0, count):
			comp = random.choice(computers)
			user = random.choice(it_users)
			props.append({'a': user, 'b': comp})
		

		session.run('UNWIND {props} AS prop MERGE (n:User {name: prop.a}) MERGE (m:Computer {name: prop.b}) MERGE (n)-[r:CanRDP]->(m)', props=props)

		props = []
		for i in xrange(0, count):
			comp = random.choice(computers)
			user = random.choice(it_users)
			props.append({'a': user, 'b': comp})
		

		session.run('UNWIND {props} AS prop MERGE (n:User {name: prop.a}) MERGE (m:Computer {name: prop.b}) MERGE (n)-[r:ExecuteDCOM]->(m)', props=props)

		props = []
		for i in xrange(0, count):
			comp = random.choice(computers)
			user = random.choice(it_groups)
			props.append({'a': user, 'b': comp})
		

		session.run('UNWIND {props} AS prop MERGE (n:Group {name: prop.a}) MERGE (m:Computer {name: prop.b}) MERGE (n)-[r:CanRDP]->(m)', props=props)

		props = []
		for i in xrange(0, count):
			comp = random.choice(computers)
			user = random.choice(it_groups)
			props.append({'a': user, 'b': comp})
		

		session.run('UNWIND {props} AS prop MERGE (n:Group {name: prop.a}) MERGE (m:Computer {name: prop.b}) MERGE (n)-[r:ExecuteDCOM]->(m)', props=props)

		props = []
		for i in xrange(0, count):
			comp = random.choice(computers)
			user = random.choice(it_users)
			props.append({'a': user, 'b': comp})
		

		session.run('UNWIND {props} AS prop MERGE (n:User {name: prop.a}) MERGE (m:Computer {name: prop.b}) MERGE (n)-[r:AllowedToDelegate]->(m)', props=props)

		props = []
		for i in xrange(0, count):
			comp = random.choice(computers)
			user = random.choice(computers)
			if (comp == user):
				continue
			props.append({'a': user, 'b': comp})
		
		session.run('UNWIND {props} AS prop MERGE (n:Computer {name: prop.a}) MERGE (m:Computer {name: prop.b}) MERGE (n)-[r:AllowedToDelegate]->(m)', props=props)

		print "Adding sessions"
		max_sessions_per_user = int(math.ceil(math.log10(self.num_nodes)))

		props = []
		for user in users:
			num_sessions = random.randrange(0, max_sessions_per_user)
			if (user in das):
				num_sessions = max(num_sessions, 1)
			
			if num_sessions == 0:
				continue
			
			for c in random.sample(computers, num_sessions):
				props.append({'a':user,'b':c})
			
			if (len(props) > 500):
				session.run('UNWIND {props} AS prop MERGE (n:User {name:prop.a}) WITH n,prop MERGE (m:Computer {name:prop.b}) WITH n,m MERGE (m)-[:HasSession]->(n)', props=props)
				props = []
		
		session.run('UNWIND {props} AS prop MERGE (n:User {name:prop.a}) WITH n,prop MERGE (m:Computer {name:prop.b}) WITH n,m MERGE (m)-[:HasSession]->(n)', props=props)

		print "Adding Domain Admin ACEs"
		props = []
		for x in computers:
			props.append({'name':x})

			if len(props) > 500:
				session.run('UNWIND {props} as prop MATCH (n:Computer {name:prop.name}) WITH n MATCH (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH m,n MERGE (m)-[r:GenericAll {isacl:true}]->(n)', props=props)
				props = []
		
		session.run('UNWIND {props} as prop MATCH (n:Computer {name:prop.name}) WITH n MATCH (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH m,n MERGE (m)-[r:GenericAll {isacl:true}]->(n)', props=props)

		for x in users:
			props.append({'name':x})

			if len(props) > 500:
				session.run('UNWIND {props} as prop MATCH (n:User {name:prop.name}) WITH n MATCH (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH m,n MERGE (m)-[r:GenericAll {isacl:true}]->(n)', props=props)
				props = []
		
		session.run('UNWIND {props} as prop MATCH (n:User {name:prop.name}) WITH n MATCH (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH m,n MERGE (m)-[r:GenericAll {isacl:true}]->(n)', props=props)

		for x in groups:
			props.append({'name':x})

			if len(props) > 500:
				session.run('UNWIND {props} as prop MATCH (n:Group {name:prop.name}) WITH n MATCH (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH m,n MERGE (m)-[r:GenericAll {isacl:true}]->(n)', props=props)
				props = []
		
		session.run('UNWIND {props} as prop MATCH (n:Group {name:prop.name}) WITH n MATCH (m:Group {name:"DOMAIN ADMINS@TESTLAB.LOCAL"}) WITH m,n MERGE (m)-[r:GenericAll {isacl:true}]->(n)', props=props)

		print "Creating OUs"
		temp_comps = computers
		random.shuffle(temp_comps)
		split_num = int(math.ceil(self.num_nodes / 10))
		split_comps = list(self.split_seq(temp_comps,split_num))
		props = []
		for i in xrange(0, 10):
			state = states[i]
			ou_comps = split_comps[i]
			ouname = "{}_COMPUTERS@TESTLAB.LOCAL".format(state)
			guid = str(uuid.uuid4())
			ou_guid_map[ouname] = guid
			for c in ou_comps:
				props.append({'compname':c,'ouguid':guid,'ouname':ouname})
				if len(props) > 500:
					session.run('UNWIND {props} as prop MERGE (n:Computer {name:prop.compname}) WITH n,prop MERGE (m:OU {guid:prop.ouguid, name:prop.ouname, blocksInheritance: false}) WITH n,m,prop MERGE (m)-[:Contains]->(n)', props=props)
					props = []
		
		session.run('UNWIND {props} as prop MERGE (n:Computer {name:prop.compname}) WITH n,prop MERGE (m:OU {guid:prop.ouguid, name:prop.ouname, blocksInheritance: false}) WITH n,m,prop MERGE (m)-[:Contains]->(n)', props=props)

		temp_users = users
		random.shuffle(temp_users)
		split_users = list(self.split_seq(temp_users,split_num))
		props = []

		for i in xrange(0, 10):
			state = states[i]
			ou_users = split_users[i]
			ouname = "{}_USERS@TESTLAB.LOCAL".format(state)
			guid = str(uuid.uuid4())
			ou_guid_map[ouname] = guid
			for c in ou_users:
				props.append({'username':c,'ouguid':guid,'ouname':ouname})
				if len(props) > 500:
					session.run(
						'UNWIND {props} as prop MERGE (n:User {name:prop.username}) WITH n,prop MERGE (m:OU {guid:prop.ouguid, name:prop.ouname, blocksInheritance: false}) WITH n,m,prop MERGE (m)-[:Contains]->(n)', props=props)
					props = []
		
		session.run('UNWIND {props} as prop MERGE (n:User {name:prop.username}) WITH n,prop MERGE (m:OU {guid:prop.ouguid, name:prop.ouname, blocksInheritance: false}) WITH n,m,prop MERGE (m)-[:Contains]->(n)', props=props)

		props = []
		for x in ou_guid_map.keys():
			guid = ou_guid_map[x]
			props.append({'b':guid})
		
		session.run('UNWIND {props} as prop MERGE (n:OU {guid:prop.b}) WITH n MERGE (m:Domain {name:"TESTLAB.LOCAL"}) WITH n,m MERGE (m)-[:Contains]->(n)', props=props)

		print "Creating GPOs"

		for i in xrange(1,20):
			gpo_name = "TESTLAB_GPO_{}@TESTLAB.LOCAL".format(i)
			guid = str(uuid.uuid4())
			session.run("MERGE (n:GPO {name:{gponame}}) SET n.guid={guid}", gponame=gpo_name, guid=guid)
			gpos.append(gpo_name)

		ou_names = ou_guid_map.keys()
		for g in gpos:
			num_links = random.randint(1,3)
			linked_ous = random.sample(ou_names,num_links)
			for l in linked_ous:
				guid = ou_guid_map[l]
				session.run("MERGE (n:GPO {name:{gponame}}) WITH n MERGE (m:OU {guid:{guid}}) WITH n,m MERGE (n)-[r:GpLink]->(m)", gponame=g, guid=guid)
		
		num_links = random.randint(1,3)
		linked_ous = random.sample(ou_names,num_links)
		for l in linked_ous:
			guid = ou_guid_map[l]
			session.run("MERGE (n:Domain {name:{gponame}}) WITH n MERGE (m:OU {guid:{guid}}) WITH n,m MERGE (n)-[r:GpLink]->(m)", gponame="TESTLAB.LOCAL", guid=guid)

		gpos.append("DEFAULT DOMAIN POLICY@TESTLAB.LOCAL")
		gpos.append("DEFAULT DOMAIN CONTROLLER POLICY@TESTLAB.LOCAL")

		acl_list = ["GenericAll"] * 10 + ["GenericWrite"] * 15 + ["WriteOwner"] * 15 + ["WriteDacl"] * 15 + ["AddMember"] * 30 + ["ForceChangePassword"] * 15 + ["ReadLAPSPassword"] * 10

		num_acl_principals = int(round(len(it_groups) * .1))
		print "Adding outbound ACLs to {} objects".format(num_acl_principals)

		acl_groups = random.sample(it_groups, num_acl_principals)
		all_principals = it_users + it_groups
		props = []
		for i in acl_groups:
			ace = random.choice(acl_list)
			ace_string = '[r:' + ace + '{isacl:true}]'
			if ace == "GenericAll" or ace == 'GenericWrite' or ace == 'WriteOwner' or ace == 'WriteDacl':
				p = random.choice(all_principals)
				p2 = random.choice(gpos)
				session.run('MERGE (n:Group {name:{group}}) MERGE (m {name:{principal}}) MERGE (n)-' + ace_string + '->(m)', group=i, principal=p)
				session.run('MERGE (n:Group {name:{group}}) MERGE (m:GPO {name:{principal}}) MERGE (n)-' + ace_string + '->(m)', group=i, principal=p2)
			elif ace == 'AddMember':
				p = random.choice(it_groups)
				session.run('MERGE (n:Group {name:{group}}) MERGE (m:Group {name:{principal}}) MERGE (n)-' + ace_string + '->(m)', group=i, principal=p)
			elif ace == 'ReadLAPSPassword':
				p = random.choice(all_principals)
				targ = random.choice(computers)
				session.run('MERGE (n {name:{principal}}) MERGE (m:Computer {name:{target}}) MERGE (n)-[r:ReadLAPSPassword]->(m)', target=targ, principal=p)
			else:
				p = random.choice(it_users)
				session.run('MERGE (n:Group {name:{group}}) MERGE (m:User {name:{principal}}) MERGE (n)-' + ace_string + '->(m)', group=i, principal=p)

		print "Marking some users as Kerberoastable"
		i = random.randint(10,20)
		i = min(i, len(it_users))
		for user in random.sample(it_users,i):
			session.run('MATCH (n:User {name:{user}}) SET n.hasspn=true', user=user)

		print "Adding unconstrained delegation to a few computers"
		i = random.randint(10,20)
		i = min(i, len(computers))
		session.run('MATCH (n:Computer {name:{user}}) SET n.unconstrainteddelegation=true', user=user)

		session.run('MATCH (n:User) SET n.owned=false')
		session.run('MATCH (n:Computer) SET n.owned=false')
		session.run('MATCH (n) SET n.domain="TESTLAB.LOCAL"')

		print "Database Generation Finished!"



if __name__ == '__main__':
	try:
		MainMenu().cmdloop()
	except KeyboardInterrupt:
		print "Exiting"
		sys.exit()

