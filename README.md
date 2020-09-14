# Cisco-Network-Survey-Tool
## What is it?
- Python script that runs a series of show commands on a set of Cisco switches and parses and outputs the following data into a .xlsx file
  - Host
  - Interface ID
  - Is enabled?
  - Is up?
  - Mode
  - Vlan(s)
  - Mac Address (only on non-trunk ports)
  - IP address
  - Subnet mask
## Imports
- Utilises [nornir](https://nornir.readthedocs.io/en/latest/) to connect and run commands to multiple Cisco switches in parallel
- Utilises [NAPALM](https://napalm.readthedocs.io/en/latest/) functions to return dictionaries containing show command information
- Utilises [ttp](https://ttp.readthedocs.io/en/latest/) to parse 'show run' output against a Jinja2 template
- Utilises [openpyxl](https://openpyxl.readthedocs.io/en/latest/) to create and add data to a .xlsx file
## Output
- For the example GNS3 topology:
![Topology](https://user-images.githubusercontent.com/38755612/93078869-20d1f200-f683-11ea-90cf-d9e89a0e8299.PNG)
- The following .xlsx file is output:

![Output](https://user-images.githubusercontent.com/38755612/93078861-1f082e80-f683-11ea-9ea7-a6b381eeed31.png)
