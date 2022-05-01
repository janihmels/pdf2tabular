


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



select_table_1 = '''SELECT Normalized_Source_9LC,'''

select_table_2 = ""
for j in statement_period_half:
  select_table_2 += (''' sum( CASE WHEN Statement_Period_Half_9LC = "{}" AND Normalized_Source_9LC <> "Pool Revenue"
                        THEN Royalty_Payable_SB ELSE NULL END) AS `{}`,'''.format(j,j))

select_table_3 = '''sum( CASE WHEN Normalized_Source_9LC <> "Pool Revenue" THEN Royalty_Payable_SB ELSE NULL END) 
                 AS `Total`
                 FROM Master WHERE Normalized_Source_9LC <> "Pool Revenue"
                 GROUP BY Normalized_Source_9LC ORDER BY `Total` DESC'''

select_table = select_table_1 + select_table_2 + select_table_3

mycursor.execute(select_table)

table = mycursor.fetchall()

pool_rev_1 = '''SELECT Normalized_Source_9LC,'''

pool_rev_2 = ""
for l in statement_period_half:
  pool_rev_2 += ('''sum( CASE WHEN Statement_Period_Half_9LC = "{}" AND Normalized_Source_9LC = "Pool Revenue"
                    THEN Royalty_Payable_SB ELSE NULL END) AS `{}`,'''.format(l,l))

pool_rev_3 = '''sum( CASE WHEN Normalized_Source_9LC = "Pool Revenue" THEN Royalty_Payable_SB ELSE NULL END) AS `Total` 
              FROM Master WHERE Normalized_Source_9LC = "Pool Revenue"
              GROUP BY Normalized_Source_9LC ORDER BY `Total` DESC'''

pool_rev = pool_rev_1 + pool_rev_2 + pool_rev_3

mycursor.execute(pool_rev)

pool_revenue =  mycursor.fetchall()

final_table = table + pool_revenue

column_names = ['Normalized Source']

for k in statement_period_half:
  column_names.append(k)
column_names.append('Total')



print(column_names)

print(final_table)