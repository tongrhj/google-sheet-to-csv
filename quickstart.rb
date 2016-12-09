require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'csv'

require 'fileutils'

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
APPLICATION_NAME = 'GOOGLE SHEET TO CSV SCRIPT'
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials',
                             "sheets.googleapis.com-ruby-quickstart.yaml")
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS_READONLY

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))

  client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(
    client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(
      base_url: OOB_URI)
    puts "Open the following URL in the browser and enter the " +
         "resulting code after authorization"
    puts url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI)
  end
  credentials
end

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

# Sample spreadsheet:
# https://docs.google.com/spreadsheets/d/{{ spreadsheet_id }}
SPREADSHEET_ID = '18O7P29_TeVexjVSiscSsLwMFrZgqAiYfvU8JBnA0rm4'
RANGE = "'MKT: A1 Referrer + Sources'!A2:E"
# Convert the columns you want from your spreadsheet into a comma seperated array of numbers
# If it's the first column, add the number one into the array
# If it's the fifth column, add the number four into the array. Don't forget to seperate the numbers with a comma!
RANGE_COLUMNS = [0, 4].sort
puts("Extracting #{RANGE} from Google Sheet")
puts("⚠️  Do they match these column indexes: #{RANGE_COLUMNS}")

# Creates unique .csv file
OUTPUT = "output_#{SecureRandom.base64(6)}.csv"
puts("🎉  Created: #{OUTPUT}")

service.get_spreadsheet_values(SPREADSHEET_ID, RANGE).values.each do |data|
  # Pick only the relevant columns from the extracted data from Google Sheets
  new_row = RANGE_COLUMNS.map {|column| data[column]}

  # Append the new_row of relevant data just created to the csv file
  CSV.open(OUTPUT, 'a+') do |csv|
    csv << new_row
  end
  puts("➕ Adding: #{new_row}")
end

puts("✅  Sheet to CSV extraction complete")
puts("Open command: open #{OUTPUT}")
