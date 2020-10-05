require 'json'
require 'rspreadsheet'

ods = Rspreadsheet.new
s = ods.create_worksheet 'Test1'

s.cells(1,1).value = 'Id'
s.cells(1,2).value = 'Actual value'
s.cells(1,3).value = 'Status'
s.cells(1,4).value = 'Impact'
s.cells(1,5).value = 'Remediation'
s.cells(1,6).value = 'Ref'

x=2
y=1

content = JSON.load(File.open('data.json').read())
content.keys.map do |sect|
    puts sect
    if content[sect]['audit'].empty?
        s[x, y] = sect.to_s
    else
        s[x,y] = "%s" % sect
        y+=1
        s[x, y] = 'vide'
        y+=1
        s[x, y] = 'results'
        y+=1
        s[x, y] = "%s" % content[sect]['impact']
        y+=1
        s[x, y] = "%s" % content[sect]['remediation']
        y+=1
        s[x, y] = "%s" % content[sect]['ref']
    end
    x += 1
    y = 1
end
ods.save('/tmp/a.ods')