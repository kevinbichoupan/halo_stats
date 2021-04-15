
####################################
# Packages
####################################

import pandas as pd
pd.options.mode.chained_assignment = None


####################################
# Data Import
####################################
print('\n\n\n')
data = pd.read_csv('halo_data.csv')
print('Data Imported')


####################################
# Data Clean
####################################

data['kd_spread'] = data['Kills'] - data['Deaths']
data['kd_ratio'] = data['Kills']/data['Deaths']
data['win_ind'] = 0
data['win_ind'].loc[data['Result'] == 'Win'] = 1



####################################
# Overall Player Stat Generation
####################################

overall_stats = pd.DataFrame(data.groupby('Player').agg(
    total_games = pd.NamedAgg(column = 'Game Number', aggfunc = 'count')
    ,total_wins = pd.NamedAgg(column = 'win_ind', aggfunc = sum)
    ,total_kills = pd.NamedAgg(column = 'Kills', aggfunc = sum)
    ,total_assists = pd.NamedAgg(column = 'Assists', aggfunc = sum)
    ,total_deaths = pd.NamedAgg(column = 'Deaths', aggfunc = sum)
)).reset_index()

overall_stats['Map'] = 'Overall'

map_stats = pd.DataFrame(data.groupby(['Player', 'Map']).agg(
    total_games = pd.NamedAgg(column = 'Game Number', aggfunc = 'count')
    ,total_wins = pd.NamedAgg(column = 'win_ind', aggfunc = sum)
    ,total_kills = pd.NamedAgg(column = 'Kills', aggfunc = sum)
    ,total_assists = pd.NamedAgg(column = 'Assists', aggfunc = sum)
    ,total_deaths = pd.NamedAgg(column = 'Deaths', aggfunc = sum)
)).reset_index()

cum_all_stats = pd.concat([overall_stats, map_stats], sort=False)

cum_all_stats['win_percentage'] = round(cum_all_stats['total_wins'] / cum_all_stats['total_games'] * 100, 1)
cum_all_stats['k/d ratio'] = round(cum_all_stats['total_kills'] / cum_all_stats['total_deaths'], 2)
cum_all_stats['k/d spread'] = cum_all_stats['total_kills'] - cum_all_stats['total_deaths']
cum_all_stats['avg kills per game'] = round(cum_all_stats['total_kills'] / cum_all_stats['total_games'], 1)
cum_all_stats['avg assists per game'] = round(cum_all_stats['total_assists'] / cum_all_stats['total_games'], 1)
cum_all_stats['avg deaths per game'] = round(cum_all_stats['total_deaths'] / cum_all_stats['total_games'], 1)


####################################
# Recent Player Stat Generation
####################################

data['Datetime'] = pd.to_datetime(data['Date'])
recent_date = data.Datetime.max()
recent_date_str = data[data['Datetime'] == recent_date].Date.max()

recent_overall_stats = pd.DataFrame(data.loc[data['Datetime'] == recent_date].groupby('Player').agg(
    total_games = pd.NamedAgg(column = 'Game Number', aggfunc = 'count')
    ,total_wins = pd.NamedAgg(column = 'win_ind', aggfunc = sum)
    ,total_kills = pd.NamedAgg(column = 'Kills', aggfunc = sum)
    ,total_assists = pd.NamedAgg(column = 'Assists', aggfunc = sum)
    ,total_deaths = pd.NamedAgg(column = 'Deaths', aggfunc = sum)
)).reset_index()

recent_overall_stats['Map'] = 'Overall'

recent_map_stats = pd.DataFrame(data.loc[data['Datetime'] == recent_date].groupby(['Player', 'Map']).agg(
    total_games = pd.NamedAgg(column = 'Game Number', aggfunc = 'count')
    ,total_wins = pd.NamedAgg(column = 'win_ind', aggfunc = sum)
    ,total_kills = pd.NamedAgg(column = 'Kills', aggfunc = sum)
    ,total_assists = pd.NamedAgg(column = 'Assists', aggfunc = sum)
    ,total_deaths = pd.NamedAgg(column = 'Deaths', aggfunc = sum)
)).reset_index()

rec_all_stats = pd.concat([recent_overall_stats, recent_map_stats], sort=False)

rec_all_stats['win_percentage'] = round(rec_all_stats['total_wins'] / rec_all_stats['total_games'] * 100, 1)
rec_all_stats['k/d ratio'] = round(rec_all_stats['total_kills'] / rec_all_stats['total_deaths'], 2)
rec_all_stats['k/d spread'] = rec_all_stats['total_kills'] - rec_all_stats['total_deaths']
rec_all_stats['avg kills per game'] = round(rec_all_stats['total_kills'] / rec_all_stats['total_games'], 1)
rec_all_stats['avg assists per game'] = round(rec_all_stats['total_assists'] / rec_all_stats['total_games'], 1)
rec_all_stats['avg deaths per game'] = round(rec_all_stats['total_deaths'] / rec_all_stats['total_games'], 1)


####################################
# Create 'THE Halo Statline.xlsx'
####################################

maps = cum_all_stats.Map.unique()

writer = pd.ExcelWriter('THE Halo Statline.xlsx', engine = 'xlsxwriter')
workbook = writer.book

for i in maps:    
	worksheet = workbook.add_worksheet(i)
	writer.sheets[i] = worksheet

	cum_total_games = cum_all_stats.loc[cum_all_stats['Map'] == i][['Player','total_games']].sort_values('total_games', ascending = False).reset_index(drop = True)
	cum_win = cum_all_stats.loc[cum_all_stats['Map'] == i][['Player','total_wins']].sort_values('total_wins', ascending = False).reset_index(drop = True)
	cum_win_percentage = cum_all_stats.loc[cum_all_stats['Map'] == i][['Player','win_percentage']].sort_values('win_percentage', ascending = False).reset_index(drop = True)
	cum_avg_stats = cum_all_stats.loc[cum_all_stats['Map'] == i][['Player', 'avg kills per game', 'avg assists per game', 'avg deaths per game']].sort_values('avg kills per game', ascending = False).reset_index(drop = True)
	cum_kdratio = cum_all_stats.loc[cum_all_stats['Map'] == i][['Player','k/d ratio']].sort_values('k/d ratio', ascending = False).reset_index(drop = True)
	cum_cumulative = cum_all_stats.loc[cum_all_stats['Map'] == i][['Player', 'total_kills', 'total_assists', 'total_deaths', 'k/d spread']].sort_values('total_kills', ascending = False).reset_index(drop = True)

	cum_dfs = [cum_total_games, cum_win, cum_win_percentage, cum_avg_stats, cum_cumulative, cum_kdratio]

	for j in cum_dfs:
	    j.index += 1

	worksheet.write(0, 0, i + ' -- Cumulative')
	worksheet.write(1, 1, 'Total Games')
	cum_total_games.to_excel(writer, sheet_name = i, startrow = 2, startcol = 0)
	worksheet.write(1, 5, 'Total Wins')
	cum_win.to_excel(writer, sheet_name = i, startrow = 2, startcol = 4)
	worksheet.write(1, 9, 'Win Percentage')
	cum_win_percentage.to_excel(writer, sheet_name = i, startrow = 2, startcol = 8)
	worksheet.write(1, 13, 'K/D Ratio')
	cum_kdratio.to_excel(writer, sheet_name = i, startrow = 2, startcol = 12)
	worksheet.write(1, 17, 'Average Stats Per Game')
	cum_avg_stats.to_excel(writer, sheet_name = i, startrow = 2, startcol = 16)
	worksheet.write(1, 23, 'Cumulative Stats')
	cum_cumulative.to_excel(writer, sheet_name = i, startrow = 2, startcol = 22)



	rec_total_games = rec_all_stats.loc[rec_all_stats['Map'] == i][['Player','total_games']].sort_values('total_games', ascending = False).reset_index(drop = True)
	rec_win = rec_all_stats.loc[rec_all_stats['Map'] == i][['Player','total_wins']].sort_values('total_wins', ascending = False).reset_index(drop = True)
	rec_win_percentage = rec_all_stats.loc[rec_all_stats['Map'] == i][['Player','win_percentage']].sort_values('win_percentage', ascending = False).reset_index(drop = True)
	rec_avg_stats = rec_all_stats.loc[rec_all_stats['Map'] == i][['Player', 'avg kills per game', 'avg assists per game', 'avg deaths per game']].sort_values('avg kills per game', ascending = False).reset_index(drop = True)
	rec_kdratio = rec_all_stats.loc[rec_all_stats['Map'] == i][['Player','k/d ratio']].sort_values('k/d ratio', ascending = False).reset_index(drop = True)
	rec_cumulative = rec_all_stats.loc[rec_all_stats['Map'] == i][['Player', 'total_kills', 'total_assists', 'total_deaths', 'k/d spread']].sort_values('total_kills', ascending = False).reset_index(drop = True)

	rec_dfs = [rec_total_games, rec_win, rec_win_percentage, rec_avg_stats, rec_cumulative, rec_kdratio]

	for j in rec_dfs:
	    j.index += 1


	vertical_buffer = len(cum_total_games) + 4

	worksheet.write(0 + vertical_buffer, 0, i + ' -- Recent Halo Night On ' + recent_date_str)
	worksheet.write(1+ vertical_buffer, 1, 'Total Games')
	rec_total_games.to_excel(writer, sheet_name = i, startrow = 2 + vertical_buffer, startcol = 0)
	worksheet.write(1+ vertical_buffer, 5, 'Total Wins')
	rec_win.to_excel(writer, sheet_name = i, startrow = 2 + vertical_buffer, startcol = 4)
	worksheet.write(1+ vertical_buffer, 9, 'Win Percentage')
	rec_win_percentage.to_excel(writer, sheet_name = i, startrow = 2 + vertical_buffer, startcol = 8)
	worksheet.write(1+ vertical_buffer, 13, 'K/D Ratio')
	rec_kdratio.to_excel(writer, sheet_name = i, startrow = 2 + vertical_buffer, startcol = 12)
	worksheet.write(1+ vertical_buffer, 17, 'Average Stats Per Game')
	rec_avg_stats.to_excel(writer, sheet_name = i, startrow = 2 + vertical_buffer, startcol = 16)
	worksheet.write(1+ vertical_buffer, 23, 'Cumulative Stats')
	rec_cumulative.to_excel(writer, sheet_name = i, startrow = 2 + vertical_buffer, startcol = 22)

writer.save()

print('THE Halo Statline.xlsx Generated')

####################################
# Packages for Emails
####################################

import configparser
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

####################################
# Define Function for Sending Emails
####################################

def send_mail(send_from,send_to,subject,text,server,port,username,password,isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("THE Halo Statline.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="THE Halo Statline.xlsx"')
    msg.attach(part)

    #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
    #SSL connection only working on Python 3+
    smtp = smtplib.SMTP(server, port)
    if isTls:
        smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to.split(","), msg.as_string())
    smtp.quit()
    print('Email sent successfully')
    

####################################
# Define Parameters for  Emails
####################################    

config = configparser.ConfigParser()
config.read('config.conf')
configs = dict(config.items('Gmail Configs'))

gmail_user = configs['gmail_user']
gmail_password = configs['gmail_password']
mailing_list = configs['mailing_list']
server = configs['server']
port = int(configs['port'])

sender = 'Master Chief'
subject = "Halo Tuesday Statline - " + recent_date_str
body = """
Soldiers,

Cheers to another successful Halo Tuesday.

Master Chief
"""

####################################
# Sending Email
####################################    

send_mail(sender, mailing_list, subject, body, server, port, gmail_user, gmail_password)

print('\n\n\n')


