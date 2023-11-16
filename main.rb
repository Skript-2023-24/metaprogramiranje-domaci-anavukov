require'./ruby.rb'


session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key("1Gu1zzZicp2FHL_hcc3pIPCt70LhtZyqbwls0C5qgdNY").worksheets[0]

tabela = GoogleSheetEnumerable.new(ws)

while true
    puts "Unesite izraz: "
    trenIzraz = gets.chomp
    if trenIzraz == "kraj"
        break
    end
    puts instance_eval trenIzraz
end
#test primeri

# tabela.each do |value, row, col|
#   puts "Red #{row}, Kolona #{col}: #{value}"
# end
# tabela['Prva kolona'][2] = 15
# incremented_values = tabela.PrvaKolona.map do |cell|
#   cell_value = cell.to_i  # Convert cell value to integer
#   cell_value + 1  # Increment the value by 1
# end
# puts "Original values: #{tabela.PrvaKolona.to_a}"
# puts "Incremented values: #{incremented_values}"


# selected_values = tabela.PrvaKolona.select do |cell|
#   cell.to_i < 5  # Convert cell value to integer and check if it's greater than 20
# end

# puts "Selected values smaller than 5: #{selected_values}"
# puts "Reduce in Prva Kolona: #{tabela.PrvaKolona.reduce(0) { |sum, value| sum + value.to_i }}"
# puts tabela.Indeks.find_row_by_value('rn1234')
