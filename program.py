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


l = len(list_of_lists)
cell_list = wks.range('U2:U%s' %(l))

for cell in cell_list:
	# cell.value = 'O_o'
	i = cell.row
	if cell.value == '':
		tc = '''=ЕСЛИ(C%s<>"";C%s;ЕСЛИ(D%s<>"";D%s;ЕСЛИ(E%s<>"";E%s;ЕСЛИ(F%s<>"";F%s;ЕСЛИ(G%s<>"";G%s;ЕСЛИ(H%s<>"";H%s;ЕСЛИ(I%s<>"";I%s;0)))))))''' %(i, i, i, i, i, i, i, i, i, i, i, i, i, i)
		wks.update_cell(i, 15, tc)
		cust_id = '''=ВПР(O%s;'Дистрибьютор'!D:F;3;0)''' %(i)
		wks.update_cell(i, 16, cust_id)
		model_id = '''=ВПР(J%s;'модель'!C:E;3;0)''' %(i)
		wks.update_cell(i, 17, model_id)
		sql = '''="INSERT INTO pda.tbl_tablets_list (IMEI_1, IMEI_2, device_number, device_model_id, owners_id, customers_cust_id, purchase_date, coment, technical_condition_id, Decommissioning_date, is_active, created, updated, updated_by) VALUES ('"&K%s&"', "&ЕСЛИ(L%s = "";"NULL";СЦЕПИТЬ("'";L%s;"'"))&", '"&M%s&"', '"&Q%s&"', 2, '"&P%s&"', getdate(), NULL, 1, NULL, 1, getdate(), getdate(), SYSTEM_USER);"''' %(i, i, i, i, i, i)
		wks.update_cell(i, 18, sql)
		sql_caps_ua = '''="INSERT INTO `device` (`device_id`, `created_at`, `updated_at`, `created_by`, `updated_by`) VALUES ('"&K%s&"', NOW(), NOW(), 'adidenko47', 'adidenko47');"''' %(i)
		wks.update_cell(i, 19, sql_caps_ua)
		sql_caps_ua_2 = '''=ЕСЛИ(L%s = "";"";("INSERT INTO `device` (`device_id`, `created_at`, `updated_at`, `created_by`, `updated_by`) VALUES ('"&L%s&"', NOW(), NOW(), 'adidenko47', 'adidenko47');"))''' %(i, i)
		wks.update_cell(i, 20, sql_caps_ua_2)
		wks.update_cell(i, 21, 'Nok!')