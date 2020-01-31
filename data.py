#!/usr/bin/env python3
"""Office Assistant Data Store."""
# FUNERAL INFO DICTIONARY LABELS #
# ------------------------------ #
srv_Labels = ['Service Number',
			  'Service Type',
			  'Service Date',
			  'Service Time',
			  'Service Day',
			  'Service Location',
			  'Celebrant',
			  'Funeral Home',
			  'F. H. Contact',
			  'Contact Phone',
			  'Organist',
			  'Cantor',
			  'Server 1',
			  'Server 2',
			  'Server 3']
			  
srv_Types = ['Full Funeral',
			 'Memorial Service',
		     'Burial Only']
		     
srv_Places = ['Mother of Sorrows',
			  'St. Joseph',
			  'Mt. Carmel',
			  'St. Joseph Cemetery',
			  'Funeral Home',
			  'Other']	# st. joseph cemetery
			  
celebrants = ['Fr. Thomas', 
			  'Fr. David', 
			  'Fr. Mulqueen', 
			  'Dec. Johnson']	

fun_Homes = ['Billman',
			 'Czup',
			 'Ducro',
			 'Fleming',
			 'Guerriero',
			 'Potti',
			 'Zaback']	   
				  				  
dpi_Labels = ['Name',
			  'Age',
			  'Date of Birth',
			  'Address',
			  'City of Birth',
			  'Date of Death',
			  'City of Death',
			  'Marital Status',
			  'Spouse',
			  'Maiden Name']
			  
dpi_MarStatus_Labels = ['Single',
						'Married',
						'Widowed',
						'Divorced',
						'Other']
				   
nok_Labels = ['Next of Kin',
			  'Relationship',
			  'Address',
			  'Phone']
			  
cem_Labels = ['Cemetery',
			  'City',
			  'Fee',
			  'Burial Date',
			  'Burial Time',
			  'Burial Day',			  
			  'Grave Location',
			  'Previous Burials',
			  'Vault Type',
			  'Urn',
			  'Military Rite']
				 
off_Labels = ['Worship Aid',
			  'Church Office']

# TIMES & DAYS FOR DROPDOWNS
times = ['7:00am',
		 '7:30am',
		 '8:00am',
		 '8:30am',
		 '9:00am',
		 '9:30am',
		 '10:00am',
		 '10:30am',
		 '11:00am',
		 '12:00pm',
		 '1:00pm',
		 '2:00pm',
		 '3:00pm',
		 '4:00pm',
		 '5:00pm']
		   
days = ['Sunday',
		'Monday',
		'Tuesday',
		'Wednesday',
		'Thursday',
		'Friday',
		'Saturday']
			
# WEB AND LOCAL PATH DICTS
fhome_links = [('Guerriero Funeral Home', "https://www.guerrierofuneralhome.com/"),
			   ('Czup Funeral Home', "https://www.czupfuneral.com/"),
			   ('Ducro Funeral Home', "https://ducro.com/"),
			   ('Potti Funeral Home', "https://www.pottifuneralhome.com/"),
			   ('Fleming Billman Funeral Home', "https://www.fleming-billman.com/"),
			   ('Our Lady of Peace', "https://olopash.org")]
			   
social_paths = {"fb": 'https://www.facebook.com/OLOPAshtabula/',
				"tw": 'https://twitter.com/OLOPAshtabula',
				"mp": 'https://myparishapp.com/',
				"ch": 'https://www.google.com',
				"yt": 'https://youtube.com/UCm_DvVqUQbbH4dTRGabUdkA'}
				
favorite_paths = {"usccb": 'http://www.usccb.org/prayer-and-worship/bereavement-and-funerals/readings-for-the-funeral-liturgy.cfm',
				  "sbo": 'https://obituaries.starbeacon.com/',
				  "doy": ''}
