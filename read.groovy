def book = GExcel.load(args[0])

def sheet = book[0]

println sheet["A1"]

println sheet.A1
sheet.A1 = "MODIFIED"
println sheet.A1

//println sheet[A1..B5]
//sheet[A1..B5].filledBy 0
//println sheet[A1..B5]

def sheet3 = book.Sheet3 // by sheet name

println sheet3.A1 as String
println sheet3.A1
println sheet3.A1
println sheet3.A2
println sheet3.A3 as int
println sheet3.A3 as Integer
println sheet3.A4 as Date
println sheet3.A5 as boolean
println sheet3.A6 as Boolean
println sheet3.B1
println sheet3.B2

