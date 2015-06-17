require 'pdf/reader'
require 'axlsx'

def process(record)
	xcl = []
	rec = record[0].split
	rec.delete(',')
	if rec.length == 9 or rec.length == 8 then 
		xcl[1] = rec[1] + " " + rec[2] + " " + rec[3]
		xcl[2] = rec[4]
		xcl[0] = rec[0]
	else
#		puts "The following record needs to be handled manually as it does not comply with regular standards."
#		print rec
		xcl[0] = rec[0]
		xcl[1] = rec[1] + " " + rec[2] + " " + rec[3]
#		puts ""	
	end
	# now handle the subjects info
	subjects = []
	for i in 1..record.length do
		start_pnta = record[i] =~ /[0-9]+\./
		if start_pnta != nil then	
			start_pntb = record[i][(start_pnta+3)..record[i].length-1] =~ /[0-9]+\./
			if start_pntb == nil then
				str_a = record[i][start_pnta..-1]
			else
				str_a = record[i][start_pnta..start_pntb-1]
			end
			str_b = ""
			if start_pntb != nil then
				str_b = record[i][(start_pntb+start_pnta+3)..record[i].length-1]
			end
			subjects[str_a[0..1].to_i] = str_a[4..-1]
			subjects[str_b[0..1].to_i] = str_b[4..-1]
		end
	end
	subjects.compact!
	for i in 0..subjects.length-1 do
#			place = subjects[i] =~ /["PP""TW""PR""OR"]\s+/
		s = subjects[i].scan(/\d+/).map(&:to_i)
#			s << subjects[i][0..place+1].gsub(/\s+/, "")
		xcl[i + 3] = s[2]
	end
	return xcl
end

def parse(page)
	triplets = []
	str = []
	in_rec = false
	page.each_line do |line|
		if (line =~ /^\s*[STFB]{1}[0-9]/) == 0 then
			in_rec = true
		elsif (line =~ /^\s*ORDN/) == 0 then
			in_rec = false
			triplets << process(str)
			str.clear
		end
		str << line if in_rec == true
	end
	return triplets
end


if ARGF.argv[0] == nil then
	puts "No arguments mentioned. Please mention name of pdf file in syntax as given below."
	puts "Syntax: ruby automate.rb \"filename.pdf\"[important] \"filename.xlsx\"[optional]"
	puts ""
	exit
else
	if (/\.pdf$/ =~ ARGF.argv[0]) != nil then
		pdfName = ARGF.argv[0]
	else
		pdfName << '.pdf'
	end
	if ARGF.argv[1] == nil then
		excelName = pdfName[0..-5] + '.xlsx'
	else
		if (/\.xlsx$/ =~ ARGF.argv[1]) != nil then
			pdfName = ARGF.argv[1]
		else
			excelName << '.pdf'
		end
	end
end
puts "**********Automation tool to convert result pdfs to excel spreadsheets**********"
puts "Note: Some records may not be according to the general standard hence the third "
puts "      column might get affected. Only student names and mother's names may get "
puts "      affected. The rest information will be put in accurately."
puts "      You may need to handle special cases separately.\n\n"
print "Reading from pdf...                                                        0%"
total = i = 0
Axlsx::Package.new do | p |
	p.workbook do | wb | 
		wb.add_worksheet do | sheet |
			pdf = PDF::Reader.new(pdfName)
			total = pdf.page_count
			pdf.pages.each do |page|
				x = parse page.text
#				print x
				x.each do |xs|
					sheet.add_row xs if xs != nil
				end
				i += 1 
				print "\rReading from pdf...                                                         #{i*100/total}%"
			end
			puts ""
			puts "Writing out to excel file..."
		end
	end
	p.serialize excelName
end
puts "Done!"
