import gspread
# тут описание как пользоваться https://github.com/burnash/gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('ReadGoogleSheets-9599725075e2.json', scope)

gc = gspread.authorize(credentials)

sht2 = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dNZbpsXAZJPgvLf137pgxucSZE4eBWYAGRTBrT0AsqM/edit#gid=1428639204')
# sht1 = gc.open_by_key('1RGsfF61l6sKke5MaKc4l2PaxFlYW4jvy43KrKtAR3co')
sht1 = gc.open_by_key('AIzaSyBJb1YVVzgAJgjIYt7ALJf406aqFSZ2i2M')

# Open a worksheet from spreadsheet with one shot
# wks = gc.open('test').sheet1
wks = sht2.sheet1


list_of_lists = wks.get_all_values()
for i in range(3,len(list_of_lists)+1):
	tc = '''=ЕСЛИ(C%s<>"";C%s;ЕСЛИ(D%s<>"";D%s;ЕСЛИ(E%s<>"";E%s;ЕСЛИ(F%s<>"";F%s;ЕСЛИ(G%s<>"";G%s;ЕСЛИ(H%s<>"";H%s;ЕСЛИ(I%s<>"";I%s;0)))))))''' %(i, i, i, i, i, i, i, i, i, i, i, i, i, i)
	wks.update_cell(i, 15, tc)
	cust_id = '''=ВПР(O%s;'Дистрибьютор'!D:F;3;0)''' %(i)
	wks.update_cell(i, 16, cust_id)
	model_id = '''=ВПР(J%s;'модель'!C:E;3;0)''' %(i)
	wks.update_cell(i, 17, model_id)
	sql = '''="INSERT INTO `pda_tbl_tablets_list` (`IMEI_1`, `IMEI_2`, `device_number`, `device_model_id`, `owners_id`, `customers_cust_id`, `purchase_date`, `coment`, `technical_condition_id`, `Decommissioning_date`, `is_active`, `created`, `updated`, `updated_by`) VALUES ('"&K%s&"', "&ЕСЛИ(L%s = "";"NULL";СЦЕПИТЬ("'";L%s;"'"))&", '"&M%s&"', '"&Q%s&"', 2, '"&P%s&"', curdate(), NULL, 1, NULL, 1, NOW(), NOW(), 1);"''' %(i, i, i, i, i, i)
	wks.update_cell(i, 18, sql)
	sql_caps_ua = '''="INSERT INTO `device` (`device_id`, `created_at`, `updated_at`, `created_by`, `updated_by`) VALUES ('"&%s&"', NOW(), NOW(), 'adidenko47', 'adidenko47');"''' %(i)
	wks.update_cell(i, 19, sql_caps_ua)
	sql_caps_ua_2 = '''=ЕСЛИ(L%s = "";"";("INSERT INTO `device` (`device_id`, `created_at`, `updated_at`, `created_by`, `updated_by`) VALUES ('"&L%s&"', NOW(), NOW(), 'adidenko47', 'adidenko47');"))''' %(i, i)
	wks.update_cell(i, 20, sql_caps_ua_2)

# i=1
# for l in list_of_lists:
# 	print(l[17])
# 	#wks.update_cell(i, 22, 'ok!')  # запись значения в столбик
# 	i +=1



# cell_list = wks.findall("ok!")
# for x in cell_list:
# 	# print(x)
# 	# print("Found something at R%sC%s" % (x.row, x.col))
# 	print(x.row)
# 	values_list = wks.row_values(x.row)
# 	print(values_list[10])

# 	s = '''="INSERT INTO `pda_tbl_tablets_list` (`IMEI_1`, `IMEI_2`, `device_number`, `device_model_id`, `owners_id`, `customers_cust_id`, `purchase_date`, `coment`, `technical_condition_id`, `Decommissioning_date`, `is_active`, `created`, `updated`, `updated_by`) VALUES ('"&K%s&"', "&ЕСЛИ(L%s = "";"NULL";СЦЕПИТЬ("'";L%s;"'"))&", '"&M%s&"', '"&Q%s&"', 2, '"&P%s&"', curdate(), NULL, 1, NULL, 1, NOW(), NOW(), 1);"''' %(x.row, x.row, x.row, x.row, x.row, x.row)
# 	print(s)

# 	wks.update_cell(x.row, 21, s)
