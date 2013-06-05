#
#  Kotie, 02-03-2013
# 
# Notice about SQL update scripts needed for v1012


NB... Only run the SQL update scripts if it was never run on a database before.


Run 0_SQL_Update_2005_2006.sql first - This will do schema updates to the database (Only add to existing
					tables, no data removal)

Run 0_SQL_Update_2005_2006_01.sql first - This will prepare newly added columns for use. All data in the  
					users.owner_transfer, products.Recipe_charge_item, users.reprint
					will be set to False with this script.