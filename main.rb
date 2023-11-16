require'./ruby.rb'


session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key("1Gu1zzZicp2FHL_hcc3pIPCt70LhtZyqbwls0C5qgdNY").worksheets[0]

tabela = GoogleSheetEnumerable.new(ws)

while true
    puts "Unesite izraz"
    trenIzraz = gets.chomp
    if trenIzraz == "kraj"
        break
    end
    puts instance_eval trenIzraz
end