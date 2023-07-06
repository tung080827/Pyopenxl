
string = "'" + "Bump Visual" +"'!" + "AB+Parameters!$C$8/2"
def getstring(string:str,c1:str, c2:str):
	cell = string
	idx1 = cell.find(c1)
	cell_2 = cell[:idx1] + cell[idx1+1 :]
	idx2 = cell_2.find(c2)
	if(idx2 != -1):
		
		cell_tmpx = cell[idx1+1:idx2+1]
		if(cell_tmpx.find("+") != -1 or cell_tmpx.find("-") != -1 or cell_tmpx.find("*") != -1  or cell_tmpx.find("/") != -1):
			return 1
		else: return 0
	else:
		return 1
string = getstring(string, "!", "!")
print(string)
print("\"")
# try:
# 	idx2 = cell_2.index("!")
# 	ref = 0
# except:
# 	ref = 1
# if(ref == 0):
# 	cell_3 = cell[idx1+1:idx2+1]
# 	print(cell_3)
# 	print(cell_3.find("-"))
# 	if(cell_3.find("+") != -1 and cell_3.find("-") != -1 and cell_3.find("*") != -1 and cell_3.find("/") != -1):
# 		print("Tung")
# 	else:
# 		print("Khong ohai Tung")
