import mysql.connector
import matplotlib.pyplot as plt
import textwrap
from docxtpl import DocxTemplate, InlineImage
import jinja2

def ReadMySQL():
	mydb = mysql.connector.connect(
	host="localhost",
	user="root",
	password="123456",
	database="athlete_data"
	)

	mycursor = mydb.cursor()

	get_name = "select name from athlete_data.athlete_record where discipline = 'pencaksilat' and level = 'Trẻ' group by name having count(*) = 2;"
	mycursor.execute(get_name) #thay code ở đây

	name_list = mycursor.fetchall()
	for each_name in name_list:
		get_data = f"select level, name, sex, birth, height_cm, weight_kg, bodyfat_percent, cmj_cm, 2arm_length_cm, 1arm_length_cm, h_length_cm, grip_kg, run_10m_s, knee_chest_push_3kg, leg_force_kg, batak_pro_60s, seat_reach_cm, beep_m, beep_vo2max, record_date from athlete_data.athlete_record where name = '{each_name[0]}' and discipline = 'pencaksilat' order by record_date;"
		mycursor.execute(get_data)
		data = mycursor.fetchall()
		process_data(data)
	# get_data = f"select level, name, sex, birth, height_cm, weight_kg, bodyfat_percent, cmj_cm, 2arm_length_cm, 1arm_length_cm, h_length_cm, grip_kg, run_10m_s, knee_chest_push_3kg, leg_force_kg, batak_pro_60s, seat_reach_cm, beep_m, beep_vo2max, record_date from athlete_data.athlete_record where discipline = 'pencaksilat' order by record_date;"
	# mycursor.execute(get_data)
	# data = mycursor.fetchall()
	# process_data(data)

def process_data(data):
	before, after = data
	h1, w1, f1, cmj1, two_arm1, one_arm1, leg1, grip1, run1, knee1, leg_force1, batak1, seat1, m1, vo2max1, date1 = before[4:]
	level, name, sex, birth, h2, w2, f2, cmj2, two_arm2, one_arm2, leg2, grip2, run2, knee2, leg_force2, batak2, seat2, m2, vo2max2, date2 = after

	# data1 = [h1, w1, f1, cmj1, two_arm1, one_arm1, leg1, grip1, run1, knee1, leg_force1, batak1, seat1, m1, vo2max1, date1]
	# for i in data1:
	# 	if i is None:
	# 		i = '-'
	
	# data2 = [level, name, sex, birth, h2, w2, f2, cmj2, two_arm2, one_arm2, leg2, grip2, run2, knee2, leg_force2, batak2, seat2, m2, vo2max2, date2]
	# for i in data2:
	# 	if i is None:
	# 		i = '-'

	date1 = date1.strftime('%m-%Y')
	date2 = date2.strftime('%m-%Y')
	level = level.upper()

	leg_force=['Lực chân (kg)', 0, 1, leg_force1, leg_force2]
	grip=['Lực bóp tay thuận (cm)', 0, 1, grip1, grip2]
	cmj=['Bật nhảy đánh tay CMJ (cm)', 0, 1, cmj1, cmj2]
	knee=['Quỳ gối đẩy bóng 3kg trước ngực (m)', 0, 1, knee1, knee2]
	run=['Chạy 10m (giây)', 0, -1, run1, run2]
	seat=['Ngồi với (cm)', 0, 1, seat1, seat2]
	batak=['Phản xạ Batak Pro 60s (lần)', 0, 1, batak1, batak2]
	m=['Beep test (m)', 0, 1, m1, m2]
	vo2max=['Beep test (VO max) (ml/kg/ph)', 0, 1, vo2max1, vo2max2]

	calculating = [leg_force, grip, cmj, knee, run, seat, batak, m, vo2max]

	leg_force[1] = leg_force2 - leg_force1
	grip[1] = grip2 - grip1
	cmj[1] = cmj2 - cmj1
	knee[1] = knee2 - knee1
	run[1] = run2 - run1
	seat[1] = seat2 - seat1
	batak[1] = batak2 - batak1
	m[1] = m2 - m1
	vo2max[1] = vo2max2 - vo2max1

	draw_data = [leg_force, grip, cmj, knee, run, seat, batak, m, vo2max]
	draw_chart(name, draw_data)
	print_word(date1, date2, leg_force, grip, cmj, knee, run, seat, batak, m, vo2max, level, name, sex, birth, f1, f2, h2, w2, two_arm2, one_arm2, leg2)
	
def draw_chart(name, data):
	plt.clf()
	data_array = []
	column_name = []
	
	for i in range(len(data)):
		d = data[i][1]*data[i][2]
		s = data[i][3]+data[i][4]
		w = (2*d/s)/2*100
		data_array.append(w)
		column_name.append(data[i][0])

	column_name = [ '\n'.join(textwrap.wrap(cat, width=10)) for cat in column_name ]

	plt.rcParams['font.family']='Times New Roman'
	plt.figure(figsize=(7,3.8))
	# ax = plt.subplots()
	bars = plt.bar(column_name, data_array)
	plt.grid(axis='y')
	plt.gca().set_axisbelow(True)

	for spine in plt.gca().spines.values():
		spine.set_visible(False)

	for bar in bars:
		yval = bar.get_height()
		yval = round(yval, 2)
		if yval > 0:
			plt.text(bar.get_x() + bar.get_width()/2.0, yval, str(yval), va='bottom', ha='center')
		else:
			plt.text(bar.get_x() + bar.get_width()/2.0, yval, str(yval), va='top', ha='center')

	plt.tick_params(axis='x', labelsize=6.5)

	plt.title(name)
	plt.savefig(f'{name}.png', dpi=300)

def print_word(date1, date2, leg_force, grip, cmj, knee, run, seat, batak, m, vo2max, level, name, sex, birth, f1, f2, h2, w2, two_arm2, one_arm2, leg2):

	template = DocxTemplate('mẫu-w-pencaksilat.docx')
	
	chart_image_path = (f"{name}.png")
	chart_image = InlineImage(template, image_descriptor=chart_image_path)

	context = {
	'level': level,
	'name': name,
	'sex': sex,
	'birth': birth,
	'height':h2,
	'weight': w2,
	'record1': date1,
	'record2': date2,
	'f1': f1,	
	'f2': f2,	
	'lf1': leg_force[3],	
	'lf2': leg_force[4],	
	'lfd': leg_force[1],	
	'grip1': grip[3],	
	'grip2': grip[4],	
	'gripd': grip[1],	
	'cmj1': cmj[3],	
	'cmj2': cmj[4],	
	'cmjd': cmj[1],	
	'chest1': knee[3],	
	'chest2': knee[4],	
	'chestd': knee[1],
	'run1': run[3],
	'run2': run[4],
	'rund': run[1],
	'seat1': seat[3],
	'seat2': seat[4],
	'seatd': seat[1],
	'batak1': batak[3],
	'batak2': batak[4],
	'batakd': batak[1],
	'm1': m[3],
	'm2': m[4],
	'md': m[1],
	'vo1': vo2max[3],
	'vo2': vo2max[4],
	'vod': vo2max[1],
	'arm2': two_arm2,
	'arm1': one_arm2,
	'leg': leg2,
	'chart': chart_image
	}

	jinja_env = jinja2.Environment(autoescape=True)
	template.render(context, jinja_env)

	template.save(f'{name}_{level}_pencaksilat.docx') # chỉnh đường dẫn vào thư mục

ReadMySQL()