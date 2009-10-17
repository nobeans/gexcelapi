def book = GExcel.load(args[0])

def sheet = book[0]

println sheet["A1"]

println sheet.A1
sheet.A1 = "MODIFIED"
println sheet.A1

println sheet.A2
println sheet.AA1
println sheet.BC1
println sheet.ZZ1
println sheet.ZZZ1

//println sheet[A1..B5]
//sheet[A1..B5].filledBy 0
//println sheet[A1..B5]

def sheet3 = book.Sheet3 // by sheet name

println sheet3.A1
println sheet3.A2
println sheet3.A3.asInt()
println sheet3.A4.asDate()
println sheet3.A5.asBoolean()
println sheet3.A6.asBoolean()
println sheet3.B1
println sheet3.B2

