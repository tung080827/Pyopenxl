
def generate_bump_table(excel_path, excel_sheet, package_type, out_table_params, bump_visual_params, dummy_params, opt_sr, die_params ):

   
  
# Bump table config 
    table={
        "name": out_table_params[0],
        "location": out_table_params[1],
        "name_wsr": out_table_params[2],
        "location_wsr": out_table_params[3],
        
    }

    #---Bump map visual view parameter---#
    coordinate = {
        
        "window1": bump_visual_params[0], #Top Left of Bump map visual view
        "window2": bump_visual_params[1], #Bottom Right of Bump map visual view
        "xcoor": bump_visual_params[2], #This define row where Xaxis value can be got
        "ycoor": bump_visual_params[3] #This define row where Yaxis value can be got
    }

    #---Dummy Bump visual view parameter---#
    dummybump={
        "corner_1":{
            "window1": dummy_params[0],
            "window2": dummy_params[1],
            "xcoor": dummy_params[2],
            "ycoor": dummy_params[3]
            },
        "corner_2":{
          
            "window1": dummy_params[4],
            "window2": dummy_params[5],
            "xcoor": dummy_params[6],
            "ycoor": dummy_params[7]
        },
        "corner_3":{
         
            "window1": dummy_params[8],
            "window2": dummy_params[9],
            "xcoor": dummy_params[10],
            "ycoor": dummy_params[11]
        },
        "corner_4":{         
            "window1": dummy_params[12],
            "window2": dummy_params[13],
            "xcoor": dummy_params[14],
            "ycoor": dummy_params[15]
        }

    }


    mynotif("")
    root.update_idletasks()
    mynotif("Loading the ploc file...")
    root.update_idletasks()
    try:
        wb = load_workbook(excel_path)
        print(wb)   
    except:
        print("Wrong Ploc path or Ploc file is openning. Please recheck/close the PLOC file before generate :(")
        show_error("Wrong Ploc path or Ploc file is openning. Please recheck/close the PLOC file before generate :(")
        progress_bar(0)
        mynotif("Error")
        root.update_idletasks()
    
    # ws = wb.create_sheet('Tung')
    try:
       ws1 = wb[excel_sheet] 
    except:
        print("Sheet name doesn't exist")
        show_error("Sheet name doesn't exist")
        progress_bar(0)
        mynotif("Error")
        root.update_idletasks()
    
    

    #----- Create dummy bump at 4 corner 140x140u for advance package (CoWos)-----------#
    ymin = coordinate_to_tuple(coordinate['window1'])[0]
    xmin = coordinate_to_tuple(coordinate['window1'])[1]
    ymax = coordinate_to_tuple(coordinate['window2'])[0]
    xmax = coordinate_to_tuple(coordinate['window2'])[1]

    print(xmin,xmax)
    print(ymin,ymax)
    progress_bar(60)
    if(opt_sr == 0):
        try:
          #----- Create table from bump map-----------#
            tb_x = coordinate_to_tuple(table['location'])[1]
            tb_y = coordinate_to_tuple(table['location'])[0]

            r = tb_y + 2
            ws1[table['location']].value = table['name']
            ws1.merge_cells(table['location'] + ":" + get_column_letter(tb_x + 2) + str(tb_y))
          
            ws1[get_column_letter(tb_x) + str(tb_y + 1)].value = "X"
            ws1[get_column_letter(tb_x + 1) + str(tb_y + 1)].value = "Y"
            ws1[get_column_letter(tb_x + 2)  + str(str(tb_y + 1))].value = "Bump name"

            ws1[get_column_letter(tb_x + 5) + str(tb_y)].value = "Die Flipped by Y axis"
            ws1.merge_cells(get_column_letter(tb_x + 5) + str(tb_y) + ":" + get_column_letter(tb_x + 7) + str(tb_y))
            ws1[get_column_letter(tb_x + 5) + str(tb_y + 1)].value = "X"
            ws1[get_column_letter(tb_x + 6) + str(tb_y + 1)].value = "Y"
            ws1[get_column_letter(tb_x + 7)  + str(str(tb_y + 1))].value = "Bump name"

            ws1[get_column_letter(tb_x + 10) + str(tb_y)].value = "DIE1 = Die Flipped rotate -90 + Die1 offset"
            ws1.merge_cells(get_column_letter(tb_x + 10) + str(tb_y) + ":" + get_column_letter(tb_x + 12) + str(tb_y))
            ws1[get_column_letter(tb_x + 10) + str(tb_y + 1)].value = "X"
            ws1[get_column_letter(tb_x + 11) + str(tb_y + 1)].value = "Y"
            ws1[get_column_letter(tb_x + 12)  + str(str(tb_y + 1))].value = "Bump name"

            ws1[get_column_letter(tb_x + 15) + str(tb_y)].value = "DIE2 = Die Flipped rotate +90 + Die2 offset"
            ws1.merge_cells(get_column_letter(tb_x + 15) + str(tb_y) + ":" + get_column_letter(tb_x + 17) + str(tb_y))
            ws1[get_column_letter(tb_x + 15) + str(tb_y + 1)].value = "X"
            ws1[get_column_letter(tb_x + 16) + str(tb_y + 1)].value = "Y"
            ws1[get_column_letter(tb_x + 17)  + str(str(tb_y + 1))].value = "Bump name"

            # xwidth = float (ws1[get_column_letter(xmax) + coordinate["xcoor"]].value)
            # minxval = float (ws1[get_column_letter(xmin) + coordinate["xcoor"]].value)
            # ywidth = float (ws1[coordinate["ycoor"] + str(ymin)].value)
            # minyval = float (ws1[coordinate["ycoor"] + str(ymax)].value)
            # xwidth = ws1[get_column_letter(xmax) + coordinate["xcoor"]].value
            # minxval = ws1[get_column_letter(xmin) + coordinate["xcoor"]].value
            # ywidth = ws1[coordinate["ycoor"] + str(ymin)].value
            # minyval = ws1[coordinate["ycoor"] + str(ymax)].value
            if (package_type == 1):
                dm_bump_coor= []
                dm_cnt=0
                mynotif("")
                root.update_idletasks()
                mynotif("Generating Dummy bump...")
                root.update_idletasks()
                for dm_bump in dummybump:
                    bump = list(dummybump[dm_bump].values())
                        
                    ymin_dm = coordinate_to_tuple(bump[0])[0]
                    xmin_dm = coordinate_to_tuple(bump[0])[1]
                    ymax_dm = coordinate_to_tuple(bump[1])[0]
                    xmax_dm = coordinate_to_tuple(bump[1])[1]
                    xcoor_dm = str(bump[2])
                    ycoor_dm = str(bump[3])

                    print(xmin_dm,xmax_dm)
                    print(ymin_dm,ymax_dm)

                    for dummycol1 in range(xmin_dm, xmax_dm + 1):
                        for dummyrow1 in range(ymin_dm, ymax_dm + 1):
                            col_dm = get_column_letter(dummycol1)
                            if (ws1[col_dm + str(dummyrow1)].value != None):
                                ws1[get_column_letter(tb_x + 2)+str(r)].value =  ws1[col_dm+ str(dummyrow1)].value
                                # print(col_l + " " + str(coordinate['xcoor']))
                                ws1[get_column_letter(tb_x)+str(r)].value = ws1[col_dm + xcoor_dm].value
                                # print(coordinate['ycoor'] + " " + str(dummyrow1)) 
                                ws1[get_column_letter(tb_x + 1)+str(r)].value = ws1[ycoor_dm + str(dummyrow1)].value
                                r = r + 1
                                coor = col_dm + str(dummyrow1)
                                dm_bump_coor.append(coor)
                                dm_cnt += 1

                #---------Create Die bump exclued dummy bump at 4 corner-----------#

                match = 0
                mynotif("")
                root.update_idletasks()
                mynotif("Generating Die bump...")
                root.update_idletasks()
                for col in range(xmin, xmax + 1):
                    for row in range(ymin, ymax + 1):       
                        col_l = get_column_letter(col)
                        #print(col_l)
                        i = 0 
                        while(i < len(dm_bump_coor)):
                            xy = col_l + str(row)
                            if(xy ==  dm_bump_coor[i]):
                                match = 1
                            else:
                                match = 0
                            if(match == 1):
                                break
                            i += 1
                        if (match == 0 and ws1[col_l + str(row)].value != None):
                            ws1[get_column_letter(tb_x + 2)+str(r)].value =  ws1[col_l+ str(row)].value
                            print(col_l + " " + str(coordinate['xcoor']))
                            ws1[get_column_letter(tb_x)+str(r)].value = ws1[col_l + str(coordinate['xcoor'])].value
                            print(coordinate['ycoor'] + " " + str(row)) 
                            ws1[get_column_letter(tb_x + 1)+str(r)].value = ws1[coordinate['ycoor'] + str(row)].value
                            r = r + 1
            else:
                process_notify("Generating Die bump...")
                for col in range(xmin, xmax + 1):
                        for row in range(ymin , ymax + 1):       
                            col_l = get_column_letter(col)
                            #print(col_l)
                            if (ws1[col_l + str(row)].value != None):
                                # ws1[get_column_letter(tb_x + 2)+str(r)].value =  ws1[col_l+ str(row)].value
                                ws1[get_column_letter(tb_x + 2)+str(r)].value =  f"={col_l+ str(row)}"
                                print(col_l + " " + str(coordinate['xcoor']))
                                ws1[get_column_letter(tb_x)+str(r)].value = ws1[col_l + str(coordinate['xcoor'])].value
                                print(coordinate['ycoor'] + " " + str(row)) 
                                ws1[get_column_letter(tb_x + 1)+str(r)].value = ws1[coordinate['ycoor'] + str(row)].value


                                
                                #----------------------------flip bump map y axis---------------------------
                                # process_notify("Flipping Die by Horizontal...")

                                ws1[get_column_letter(tb_x + 7)+str(r)].value =  ws1[col_l+ str(row)].value
                                print(col_l + " " + str(coordinate['xcoor']))
                                ws1[get_column_letter(tb_x + 5)+str(r)].value = f"=({str(die_params['chip_width']).replace('=','')})-({str(ws1[col_l + str(coordinate['xcoor'])].value).replace('=','')})"
                                print(ws1[get_column_letter(tb_x + 5)+str(r)].value)
                                print(coordinate['ycoor'] + " " + str(row)) 
                                ws1[get_column_letter(tb_x + 6)+str(r)].value = ws1[coordinate['ycoor'] + str(row)].value

                                #----------------------------rotate -90 bump map after flip---------------------------
                                # process_notify("Rotate Die flipped -90 degree...")

                                if( ws1[col_l+ str(row)].value == "VSS"):
                                    ws1[get_column_letter(tb_x + 12)+str(r)].value =  ws1[col_l+ str(row)].value
                                else:
                                    ws1[get_column_letter(tb_x + 12)+str(r)].value =  "DIE3_" + str(ws1[col_l+ str(row)].value)
                                print(col_l + " " + str(coordinate['xcoor']))
                                ws1[get_column_letter(tb_x + 10)+str(r)].value = f"=({str(die_params['chip_height']).replace('=','')})-({str(ws1[coordinate['ycoor']+str(row)].value).replace('=','')})+({str(die_params['die1_xoffset'])})"
                                print(coordinate['ycoor'] + " " + str(row)) 
                                ws1[get_column_letter(tb_x + 11)+str(r)].value = f"=({str(die_params['chip_width']).replace('=','')})-({str(ws1[col_l + str(coordinate['xcoor'])].value).replace('=','')})+({str(die_params['die1_yoffset'])})"

                                #---------------------------rotate 90 bump map after flip -----------------------------
                                # process_notify("Rotate Die flipped 90 degree...")
                                if( ws1[col_l+ str(row)].value == "VSS"):
                                    ws1[get_column_letter(tb_x + 17)+str(r)].value = ws1[col_l+ str(row)].value
                                else:
                                    ws1[get_column_letter(tb_x + 17)+str(r)].value = "DIE7_" + str(ws1[col_l+ str(row)].value) 

                                print(col_l + " " + str(coordinate['xcoor']))
                                ws1[get_column_letter(tb_x + 15)+str(r)].value = f"=({str(ws1[coordinate['ycoor'] + str(row)].value).replace('=','')})+({str(die_params['die2_xoffset']).replace('=','')})"
                                print(coordinate['ycoor'] + " " + str(row)) 
                                ws1[get_column_letter(tb_x + 16)+str(r)].value = f"=({str(ws1[col_l + str(coordinate['xcoor'])].value).replace('=','')})+({str(die_params['die2_yoffset']).replace('=','')})"

                                r = r + 1
                # tab = Table(displayName="Table1", ref="O65:Q500")
                # ws1.add_table(tab)
            
            progress_bar(80)    
            wb.save(excel_path)
            progress_bar(100)
            mynotif("Generated")
            popup("PLOC generated successful!!!")
            mynotif("")
        except (ValueError):
            print ("loi roi")
            show_error("Wrong input, Please check and regenerate")
            progress_bar(0)
            mynotif("Error")
            root.update_idletasks()
        except:
            print('Loi quan que` gi` za^y')
            
            show_error("Wrong input, Please check and regenerate")
            progress_bar(0)
            mynotif("Error")
            root.update_idletasks()
            
            
    elif(opt_sr == 1):

        tb_x = coordinate_to_tuple(table['location'])[1]
        tb_y = coordinate_to_tuple(table['location'])[0]
        tb_x2 = coordinate_to_tuple(table['location_wsr'])[1]
        tb_y2 = coordinate_to_tuple(table['location_wsr'])[0]
        r = tb_y + 2
        r2 = tb_y2 + 2

        ws1[table['location']].value = table['name']
        ws1[table['location_wsr']].value = table['name_wsr']
        # ws1.merge_cells(table['xcol'] + str(table['begin']) + ":" + table['bumpcol'] + str(table['begin']))
        # print(table['xcol'] + str(table['begin']) + ":" + table['bumpcol'] + str(table['begin']))
        ws1[get_column_letter(tb_x) + str(tb_y + 1)].value = "X"
        ws1[get_column_letter(tb_x2) + str(tb_y2 + 1)].value = "X"
        ws1[get_column_letter(tb_x + 1) + str(tb_y + 1)].value = "Y"
        ws1[get_column_letter(tb_x2 + 1) + str(tb_y2 + 1)].value = "Y"
        ws1[get_column_letter(tb_x + 2)  + str(str(tb_y + 1))].value = "Bump name"
        ws1[get_column_letter(tb_x2 + 2)  + str(str(tb_y2 + 1))].value = "Bump name"
        if (package_type == 1):
            dm_bump_coor= []
            dm_cnt=0
            mynotif("")
            root.update_idletasks()
            mynotif("Generating Dummy bump...")
            root.update_idletasks()
            for dm_bump in dummybump:
                bump = list(dummybump[dm_bump].values())
                    
                ymin_dm = coordinate_to_tuple(bump[0])[0]
                xmin_dm = coordinate_to_tuple(bump[0])[1]
                ymax_dm = coordinate_to_tuple(bump[1])[0]
                xmax_dm = coordinate_to_tuple(bump[1])[1]
                xcoor_dm = str(bump[2])
                ycoor_dm = str(bump[3])

                print(xmin_dm,xmax_dm)
                print(ymin_dm,ymax_dm)

                for dummycol1 in range(xmin_dm, xmax_dm + 1):
                    for dummyrow1 in range(ymin_dm, ymax_dm + 1):
                        col_dm = get_column_letter(dummycol1)
                        if (ws1[col_dm + str(dummyrow1)].value != None):
                            ws1[get_column_letter(tb_x + 2)+str(r)].value =  ws1[col_dm+ str(dummyrow1)].value
                            ws1[get_column_letter(tb_x2 + 2)+str(r2)].value =  ws1[col_dm+ str(dummyrow1)].value 
                      
                            ws1[get_column_letter(tb_x)+str(r)].value = ws1[col_dm + xcoor_dm].value
                           
                            ws1[get_column_letter(tb_x + 1)+str(r)].value = ws1[ycoor_dm + str(dummyrow1)].value
                            
                            r = r + 1
                            r2 = r2 + 1
                            coor = col_dm + str(dummyrow1)
                            dm_bump_coor.append(coor)
                            dm_cnt += 1

            #---------Create Die bump exclued dummy bump at 4 corner-----------#

            match = 0
            mynotif("")
            root.update_idletasks()
            mynotif("Generating Die bump...")
            root.update_idletasks()
            for col in range(xmin, xmax + 1):
                for row in range(ymin, ymax + 1):       
                    col_l = get_column_letter(col)
                    #print(col_l)
                    i = 0 
                    while(i < len(dm_bump_coor)):
                        xy = col_l + str(row)
                        if(xy ==  dm_bump_coor[i]):
                            match = 1
                        else:
                            match = 0
                        if(match == 1):
                            break
                        i += 1
                    if (match == 0 and ws1[col_l + str(row)].value != None):
                        ws1[get_column_letter(tb_x + 2)+str(r)].value =  ws1[col_l+ str(row)].value
                        ws1[get_column_letter(tb_x2 + 2)+str(r2)].value =  ws1[col_l+ str(row)].value
                        print(col_l + " " + str(coordinate['xcoor']))
                        ws1[get_column_letter(tb_x)+str(r)].value = ws1[col_l + str(coordinate['xcoor'])].value
                       
                        print(coordinate['ycoor'] + " " + str(row)) 
                        ws1[get_column_letter(tb_x + 1)+str(r)].value = ws1[coordinate['ycoor'] + str(row)].value
                        
                        r = r + 1
                        r2 = r2 + 1
        else:
            mynotif("")
            root.update_idletasks()
            mynotif("Generating Die bump...")
            root.update_idletasks()
            for col in range(xmin, xmax + 1):
                    for row in range(ymin , ymax + 1):       
                        col_l = get_column_letter(col)
                        #print(col_l)
                        if (ws1[col_l + str(row)].value != None):
                            ws1[get_column_letter(tb_x + 2)+str(r)].value =  ws1[col_l+ str(row)].value
                            ws1[get_column_letter(tb_x2 + 2)+str(r2)].value =  ws1[col_l+ str(row)].value
                            print(col_l + " " + str(coordinate['xcoor']))
                            ws1[get_column_letter(tb_x)+str(r)].value = ws1[col_l + str(coordinate['xcoor'])].value
                            
                            print(coordinate['ycoor'] + " " + str(row)) 
                            ws1[get_column_letter(tb_x + 1)+str(r)].value = ws1[coordinate['ycoor'] + str(row)].value
                           
                            r = r + 1
                            r2 = r2 + 1