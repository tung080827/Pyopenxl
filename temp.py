# bumplist:list = ['VDD', 'VCCIO', 'VSS', 'BP_TXDATASB', 'BP_TXDATA[5]', 'BP_TXDATA[4]', 'BP_TXDATA[1]', 'BP_TXDATA[0]', 'BP_TXDATA[7]', 
#                  'BP_TXDATA[6]', 'BP_TXDATA[3]', 'BP_TXDATA[2]', 'BP_TXCKSB', 'BP_TXCKN', 'BP_TXCKP', 'BP_TXVLD', 'BP_TXTRK', 'BP_TXDATA[9]', 
#                  'BP_TXDATA[8]', 'BP_TXDATA[13]', 'BP_TXDATA[12]', 'VCCAON', 'BP_TXDATA[11]', 'BP_TXDATA[10]', 'BP_TXDATA[15]', 'BP_TXDATA[14]',
#                  'BP_RXDATA[10]', 'BP_RXDATA[11]', 'BP_RXDATA[14]', 'BP_RXDATA[15]', 'BP_RXDATA[8]', 'BP_RXDATA[9]', 'BP_RXDATA[12]', 'BP_RXDATA[13]',
#                  'BP_RXCKSB', 'BP_RXCKP', 'BP_RXCKN', 'BP_RXTRK', 'BP_RXVLD', 'BP_RXDATA[6]', 'BP_RXDATA[7]', 'BP_RXDATA[2]', 'BP_RXDATA[3]', 'BP_RXDATASB',
#                  'BP_RXDATA[4]', 'BP_RXDATA[5]', 'BP_RXDATA[0]', 'BP_RXDATA[1]']

# def getstring(string: str,c1: str, c2: str):
#     cell = string
#     idx1 = cell.find(c1)
#     idx2 = cell.find(c2)
#     if(idx1 == -1 or idx2 == -1):
#         return None,None, None
#     else:
#         str_wo_c = cell[idx1+1:idx2]
#         str_w_c = cell[idx1:idx2+1]
#         str_cut = cell[:idx1]
#         return str_wo_c,str_w_c, str_cut
# def get_indexsubstring(string:str,sub:str):
#     count_er=0
#     start_index=0
#     idx:list=[]
#     for i in range(len(string)):
#         j = string.find(sub,start_index)
#         if(j!=-1):
#             start_index = j+1
#             count_er+=1
#             idx.append(j)
#         print("Total occurrences are: ", count_er)
#         print("index: ",idx)
#     return idx
# def get_lastsub(string:str, singlechar:str):
#     idx_ls = get_indexsubstring(string, singlechar)
#     return string[idx_ls[len(idx_ls)-1]:]

# power_list = ['VDD', 'VCCIO', 'VCCAON', 'VSS']
# buschar ="[]"
# def get_bus(bumplist: list, power_list:list, buschar:str):
#     buscharls = list(buschar)
#     netdict : dict ={}
#     for pwr in power_list:
#         bumplist.remove(pwr)
#     for net in bumplist:
#         s = getstring(net,buscharls[0], buscharls[1])
#         if s[0] == None:
#             if net not in power_list:
#                 netdict.__setitem__(net,1)
#     keysList = list(netdict.keys())
#     for key in keysList:
#         bumplist.remove(key)
#     print(bumplist)

#     while bumplist:
#         s = getstring(bumplist[0],buscharls[0], buscharls[1])
#         cnt = 0
#         templist = []
#         for net in bumplist:
#             if str(net).find(s[2]) != -1:
#                 cnt +=1
#                 templist.append(net)
#         netdict.__setitem__(s[2],cnt)
#         for ls in templist:
#             bumplist.remove(ls)
#     print(netdict)
#     return netdict
# get_bus(bumplist, power_list, buschar)
# # for net in bumplist:
# #     s = getstring(net,buscharls[0], buscharls[1])
# #     bus = 1
# #     if s[0] != None:
# #         for l1 in bumplist:
# #             if str(l1).find(s[2]) != -1:
# #                 bus +=1
# #                 bumplist.remove(l1)
# #             netdict.__setitem__(s[2], bus)

# # print(netdict)

#####################################################

import random

def generate_color():
    color = '#{:02x}{:02x}{:02x}'.format(*map(lambda x: random.randint(0, 255), range(3)))
    return color
print(generate_color())