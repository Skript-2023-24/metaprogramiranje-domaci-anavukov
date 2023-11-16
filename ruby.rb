require "google_drive"
require "forwardable"

class GoogleSheetEnumerable
  include Enumerable

  MergeCell = Struct.new(:start_row, :start_col, :end_row, :end_col)

  def initialize(worksheet)
    @ws = worksheet
    @merged_cells = [
      MergeCell.new(19, 6, 20, 6),
      MergeCell.new(18, 5, 18, 6),
      MergeCell.new(18, 6, 18, 7)
    ]
    @headers = extract_headers
    define_column_methods
  end

  def extract_headers
    @ws.rows.each_with_object({}).with_index do |(row, headers), row_index|
      row.each_with_index do |cell, col_index|
        normalized_cell = cell.to_s.strip.downcase
        headers[normalized_cell] = col_index if normalized_cell != ''
      end
    end
  end

  class Column
    attr_reader :header_index

    def initialize(worksheet, col_index, header_index)
      @worksheet = worksheet
      @col_index = col_index
      @header_index = header_index
    end

    def [](row_index)
      @worksheet[@header_index + row_index, @col_index]
    end

    def []=(row_index, value)
      @worksheet[@header_index + row_index, @col_index] = value
      @worksheet.save
    end
    def to_a(ignore_totals: false)
      start_row = @header_index + 1
      @worksheet.rows[start_row..-1].reject do |row|
        ignore_totals && row.any? { |cell| cell.match?(/total|subtotal/i) }
      end.map { |row| row[@col_index - 1] }
    end
  
    def sum
      to_a(ignore_totals: true)
        .map { |val| parse_number(val) }
        .compact
        .reduce(0.0, :+)
    end
  
    def avg
      numeric_values = to_a(ignore_totals: true)
                        .map { |val| parse_number(val) }
                        .compact
      numeric_values.empty? ? 0.0 : numeric_values.reduce(:+) / numeric_values.size
    end
  
    def parse_number(string)
      string.to_s.strip.match?(/\A-?\d+(\.\d+)?\z/) ? string.to_f : nil
    end
    def values
      to_a
    end

    def select
      to_a.select { |cell| yield cell }
    end

    def reduce(initial)
      to_a.reduce(initial) { |acc, cell| yield acc, cell }
    end
    def map
      to_a.map { |cell| yield cell }
    end
    def find_row_by_value(value)
      @worksheet.rows.each_with_index do |row, index|
        next if index.zero?
        cell = row[@col_index - 1].to_s.strip
        return row if cell == value
      end
      nil
    end
  end
  

  def [](header)
    header = header.to_s.strip.downcase
    col_index = @headers[header]
    return nil if col_index.nil?

    header_row_index = @ws.rows.index { |row| row[col_index].to_s.strip.downcase == header }&.+(1)
    return nil if header_row_index.nil?

    Column.new(@ws, col_index + 1, header_row_index)
  end

  def row(row_num)
    (1..@ws.num_cols).map { |col_num| @ws[row_num, col_num] }
  end
  def each
    1.upto(@ws.num_rows) do |row_index|
      row = @ws.rows[row_index - 1]
      next if row_empty?(row)
  
      row.each_with_index do |cell_value, col_index|
        next if cell_value.to_s.strip.empty?  
        next if merged_cell?(row_index, col_index + 1) && !merged_cell_start?(row_index, col_index + 1)
        
        
        yield cell_value, row_index, col_index + 1
      end
    end
  end
  
  def row_empty?(row)
    row.all? { |cell| cell.to_s.strip.empty? }
  end
 
  private

  def define_column_methods
    @headers.each_key do |header|
      method_name = header.split.map(&:capitalize).join
      define_singleton_method(method_name) { self[header] }
    end
  end

  def merged_cell_start?(row, col)
    @merged_cells.any? { |merged_cell| merged_cell.start_row == row && merged_cell.start_col == col }
  end

  def merged_cell?(row, col)
    @merged_cells.any? do |merged_cell|
      (merged_cell.start_row..merged_cell.end_row).cover?(row) &&
        (merged_cell.start_col..merged_cell.end_col).cover?(col) &&
        !(merged_cell.start_row == row && merged_cell.start_col == col)
    end
  end
 
end