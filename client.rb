require File.join(File.dirname(__FILE__), 'exasol')
require 'spreadsheet'
require 'yaml'

@no_subids_in_has_offers_report = 0
@no_subid_in_database = 0
@different_provisions = 0
@diferent_affiliate_network_status = 0
@different_affiliate_offer_id = 0
@rejected = 0
@approved = 0
@pending = 0
@different_status = 0
@testoffer = 0

@ams_provision = 0
@hasoffers_approved_provision = 0
@hasoffers_rejected_provision = 0
@hasoffers_pending_provision = 0

config = YAML.load_file("config.yaml")
@login = config["config"]["login"]
@password = config["config"]["password"]

#Create result file

result_excel = Spreadsheet::Workbook.new
sheet1 = result_excel.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{id offer_id offer affiliate_id affiliate date_time status status_message payout revenue sale_amount conversion_ip transaction_id affiliate_sub_id1 sp_status affiliate_network_status affiliate_offer_id tar_provision landing_page_id}

row_counter = 1

@connection = Exasol.new(@login, @password)
@connection.connect

#Read Excel File
Spreadsheet.open('ho_stats.xls') do |book|
  book.worksheet('Sheet1').each do |row|
    break if row[0].nil?
      next if row[13] == "affiliate_sub_id1"
      subid = row[13]
      if subid.nil?
        excel_row = sheet1.row(row_counter)
        excel_row[0] = row[0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]
        excel_row[5] = row[5]
        excel_row[6] = row[6]
        excel_row[7] = row[7]
        excel_row[8] = row[8]
        excel_row[9] = row[9]
        excel_row[10] = row[10]
        excel_row[11] = row[11]
        excel_row[12] = row[12]
        excel_row[13] = "no subid"
        excel_row[14] = "null"
        excel_row[15] = "null"
        excel_row[16] = "null"
        excel_row[17] = "null"
        excel_row[18] = "null"
        @no_subids_in_has_offers_report += 1
          if row[6] == "approved"
            @approved += 1
            @hasoffers_approved_provision += row[8]
          elsif row[6] == "rejected"
            @rejected += 1
            @hasoffers_rejected_provision += row[8]
          elsif row[6] == "pending"
            @pending += 1
            @hasoffers_pending_provision += row[8]
          else
            @different_status += 1
          end
      end
      
      puts "subid: " + "#{subid}"
 
      if subid.length > 32 && subid.length == 36
        sub = subid
      elsif subid == "testoffer"
        sub = subid
        @testoffer += 1
      else
        sub = subid.insert(8, '-').insert(13, '-').insert(18, '-').insert(23, '-')
      end

      query = "select tar.sp_status, tar.affiliate_network_status, lp.affiliate_offer_id, tar.provision, lp.id from ids.track_affiliate_responses as tar right join ids.track_offer_clicks as toc on tar.subid = toc.subid join cms.landing_pages as lp on toc.landing_page_id = lp.id where toc.subid = '#{sub}'"
      @connection.do_query(query)
      result = @connection.print_result_array
      if result.empty?
        excel_row = sheet1.row(row_counter)
        excel_row[0] = row[0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]
        excel_row[5] = row[5]
        excel_row[6] = row[6]
        excel_row[7] = row[7]
        excel_row[8] = row[8]
        excel_row[9] = row[9]
        excel_row[10] = row[10]
        excel_row[11] = row[11]
        excel_row[12] = row[12]
        excel_row[13] = row[13]
        excel_row[14] = "null"
        excel_row[15] = "null"
        excel_row[16] = "null"
        excel_row[17] = "null"
        excel_row[18] = "null"
        @no_subid_in_database += 1
          if row[6] == 'approved'
            @approved += 1
            @hasoffers_approved_provision += row[8] 
          elsif row[6] == 'rejected'
            @rejected += 1
            @hasoffers_rejected_provision += row[8]
          elsif row[6] == 'pending'
            @pending += 1
            @hasoffers_pending_provision += row[8]
          else
            @different_status += 1
          end
        puts "subid: " + "#{row[13]}"
      else
        excel_row = sheet1.row(row_counter)
        excel_row[0] = row[0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]
        excel_row[5] = row[5]
        excel_row[6] = row[6]
        excel_row[7] = row[7]
        excel_row[8] = row[8]
        excel_row[9] = row[9]
        excel_row[10] = row[10]
        excel_row[11] = row[11]
        excel_row[12] = row[12]
        excel_row[13] = row[13]
        excel_row[14] = result[0][0]
        excel_row[15] = result[0][1]
        excel_row[16] = result[0][2]
        excel_row[17] = result[0][3]
        excel_row[18] = result[0][4]
        puts "subid: " + "#{row[13]}"
        
          if result[0][3] != row[9]
            @different_provisions += 1
          elsif result[0][1] != row[6]
            @diferent_affiliate_network_status += 1
          elsif result[0][2] != row[1]
            @different_affiliate_offer_id += 1
          end

          if row[6] == 'approved'
            @approved += 1
            @hasoffers_approved_provision += row[8]
          elsif row[6] == 'rejected'
            @rejected += 1
            @hasoffers_rejected_provision += row[8]
          elsif row[6] == 'pending'
            @pending += 1
            @hasoffers_pending_provision += row[8]
          else
            @different_status += 1
          end

          @ams_provision += result[0][3] 

      end    
      row_counter += 1
  end
end

@connection.disconnect
result_excel.write 'result.xls'

puts "Number of conversions which have no AMS subids: #{@no_subids_in_has_offers_report}"
puts "Number of different provisions: #{@different_provisions}"
puts "Number of different affiliate network status: #{@diferent_affiliate_network_status}"
puts "Number of different offer ids: #{@different_affiliate_offer_id}"
puts "Number of approved: #{@approved}"
puts "Number of rejected: #{@rejected}"
puts "Number of pending: #{@pending}"
puts "Number of other statuses: #{@different_status}"
puts "Number of testoffer conversions: #{@testoffer}"
puts "AMS provision: #{@ams_provision}"
puts "HasOffers approved provision: #{@hasoffers_approved_provision}"
puts "HasOffers rejected provision: #{@hasoffers_rejected_provision}"
puts "HasOffers pending provision: #{@hasoffers_pending_provision}"

# Create a new file and write to it  
File.open('has_offers_result.txt', 'w') do |f|
  f.puts "Number of conversions which have no AMS subids: " + "#{@no_subids_in_has_offers_report}" + "\n"
  f.puts "Number of different provisions: " + "#{@different_provisions}" + "\n"
  f.puts "Number of different affiliate network status:" + "#{@diferent_affiliate_network_status}" + "\n"
  f.puts "Number of different offer ids:" + "#{@different_affiliate_offer_id}" + "\n"
  f.puts "Number of approved: " + "#{@approved}" + "\n"
  f.puts "Number of rejected: " + "#{@rejected}" + "\n"
  f.puts "Number of pending: " + "#{@pending}" + "\n"
  f.puts "Number of other statuses: " + "#{@different_status}" + "\n"
  f.puts "Number of testoffer conversions: " + "#{@testoffer}" + "\n"
  f.puts "AMS provision: " + "#{@ams_provision}" + "\n"
  f.puts "HasOffers approved provision: " + "#{@hasoffers_approved_provision}" + "\n"
  f.puts "HasOffers rejected provision: " + "#{@hasoffers_rejected_provision}" + "\n"
  f.puts "HasOffers pending provision: " + "#{@hasoffers_pending_provision}" + "\n"
end
