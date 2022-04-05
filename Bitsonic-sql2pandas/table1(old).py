

# WASN'T CHANGED: -------

# NO FUNCTION WAS DECLARED HERE :

import mysql.connector

mydb = mysql.connector.connect(
  host="34.65.111.142",
  user="external",
  password="musicpass",
  database="Matthew Knowles Take 2_61413e9d9767146b54b829e0"
)

mycursor = mydb.cursor(buffered=True)

find_period = "SELECT DISTINCT Statement_Period_Half_9LC FROM Master ORDER BY Statement_Period_Half_9LC"

mycursor.execute(find_period)

statement_period_half = [i[0] for i in mycursor.fetchall()]

print(statement_period_half)



select_table_1 = '''SELECT Song_Name_9LC,'''

select_table_2 = ""
for j in statement_period_half:
  select_table_2 += (''' sum( CASE WHEN Statement_Period_Half_9LC = "{}", Song_Name_9LC =! "Pool Revenue"
                        THEN Royalty_Payable_SB ELSE NULL END) AS `{}`,'''.format(j,j))

select_table_3 = '''sum( Royalty_Payable_SB) AS `Total`
                 FROM Master
                 GROUP BY Song_Name_9LC ORDER BY `Total` DESC'''

select_table = select_table_1 + select_table_2 + select_table_3

mycursor.execute(select_table)

table = mycursor.fetchall()

column_names = ['Song Title']

for k in statement_period_half:
  column_names.append(k)
column_names.append('Total')

print(table)

print(column_names)







