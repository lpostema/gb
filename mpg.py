import xml.etree.ElementTree as ET
import os, shutil
import datetime
tree = ET.parse('Meeting Plan Generator.xml')
root = tree.getroot()

today = datetime.date.today()
daysToNextWednesday = datetime.timedelta((2 - datetime.date.weekday(today)) % 7)
nextWednesday = today + daysToNextWednesday
nextWednesdayMonth = str(nextWednesday.month) if (nextWednesday.month >= 10) else '0' + str(nextWednesday.month)
nextWednesdayDay = str(nextWednesday.day) if (nextWednesday.day >=10) else '0' + str(nextWednesday.day)
wednesdaysDate = str(nextWednesday.year) + '-' + nextWednesdayMonth + '-' + nextWednesdayDay

for plugin in root.findall('plugin'):
    for filename in plugin.findall('filelist'):
        for file in filename.iter():
            if 'name' in file.attrib.keys():
                startIndex = file.attrib['name'].find('change-date')
                if startIndex != -1:
                    newString = file.attrib['name'][:startIndex]
                    newString += wednesdaysDate
                    newString += '\\Meeting Planner.pdf'
                    file.set('name', newString)
    for destination in plugin.findall('destination'):
        print(destination.tag, destination.attrib)
        if 'value' in destination.attrib.keys():
            startIndex = destination.attrib['value'].find('change-date')
            if startIndex != -1:
                newString = destination.attrib['value'][:startIndex]
                newString += wednesdaysDate
                newString += '\\Full Meeting Plan.pdf'
                destination.set('value', newString)

os.mkdir(wednesdaysDate)
tree.write(wednesdaysDate + '\Meeting Plan Generator 2.xml')
shutil.copyfile('Meeting Planner.xlsx', wednesdaysDate + '\Meeting Planner.xlsx')
