require 'roo'
require 'spreadsheet'
require 'geokit'
require 'byebug'

class FilterStore
  def copy_file(input_file, output_file, filter)
    file_exist = File.exist?(output_file)
    opened_input_file = Roo::Spreadsheet.open(input_file)
    count = 1

    if file_exist
      opened_output_file = Spreadsheet.open(output_file)
      sheet1 = opened_output_file.worksheet 0
      new_row_index = sheet1.last_row_index + 2
      sheet1.insert_row((new_row_index - 1), ["********"])
      sheet1.insert_row(new_row_index, opened_input_file.sheet(0).row(1))

      opened_input_file.each do |row|
        if row[0] == filter
          sheet1.insert_row((new_row_index + count), row)
          count += 1
        end
      end

      File.delete(output_file)
      opened_output_file.write output_file
    else
      opened_output_file = Spreadsheet::Workbook.new
      sheet1 = opened_output_file.create_worksheet(name: 'Sheet1')
      sheet1.row(0).replace opened_input_file.sheet(0).row(1)

      opened_input_file.each do |row|
        if row[0] == filter
          sheet1.row(count).replace row
          count += 1
        end
      end

      opened_output_file.write output_file
    end
  end

  def distance(point1, point2)
    Geokit::default_units = :meters

    current_location =  Geokit::LatLng.new(point1[0], point1[1])
    destination = Geokit::LatLng.new(point2[0], point2[1])

    current_location.distance_to(destination)
  end

  def merge_duplicated_point(merged_file, radius)
    file = Spreadsheet.open(merged_file)
    points = get_coordinates(merged_file)

    sheet1 = file.worksheet 0
    sheet1.each do |row|
      next if row[1].to_i == 0
      points -= [points[0]]
      
      points.each do |point|
        if (distance([row[1], row[2]], [point[0], point[1]]) <= radius)
          sheet1.insert_row(row.idx, [])
          sheet1.row(point[2]).push "merged with row #{row.idx + 1} [#{row[1]}, #{row[2]}]"
        end
      end
    end

    sleep 3

    File.delete(merged_file)
    file.write merged_file
  end

  def get_coordinates(merged_file)
    file = Spreadsheet.open(merged_file)
    coordinates = []

    sheet1 = file.worksheet 0
    sheet1.each do |row|
      next if row[1].to_i == 0
      coordinates << [row[1], row[2], row.idx]
    end
    coordinates
  end
end

customer_file = 'edeka_from_customer.xlsx'
pos_pulse_file = 'edeka_from_pospulse.xlsx'
target_file = 'edeka_stores.xlsx'
filter = 'Edeka'

file = FilterStore.new

given_files = [customer_file, pos_pulse_file]
given_files.each do |given_file|
  file.copy_file(given_file, target_file, filter)
end

file.merge_duplicated_point(target_file, 500)