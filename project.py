import pandas as pd
import os
import json
import uuid
from typing import Dict, List, Any
import warnings

warnings.filterwarnings('ignore')

# Update these paths to point to your Downloads folder
input_file_path = "/Users/murodjonyusupjonov/Downloads/raw_data.xlsx"
output_file_path = "/Users/murodjonyusupjonov/Downloads/final.xlsx"

class CommunicationDataETL:
    def __init__(self):
        self.dimensions = {
            "comm_type": [],
            "subject": [],
            "calendar": [],
            "audio": [],
            "video": [],
            "transcript": [],
            "user": []
        }
        self.fact_records = []
        self.bridge_records = []
        
        self.lookup_tables = {
            "comm_type": {},
            "subject": {},
            "calendar": {},
            "audio": {},
            "video": {},
            "transcript": {},
            "user": {}
        }
        
        self.id_counters = {
            "comm_type": 1,
            "subject": 1,
            "calendar": 1,
            "audio": 1,
            "video": 1,
            "transcript": 1,
            "user": 1  
        }

    def load_and_parse_data(self, file_path: str) -> List[Dict]:
        """Load and parse data from the specified file."""
        print("Step 1: Loading and parsing data...")

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        if file_path.endswith(".csv"):
            data_frame = pd.read_csv(file_path)
        else:
            data_frame = pd.read_excel(file_path, engine="openpyxl")

        print(f"Loaded {len(data_frame)} rows.")

        if "raw_content" not in data_frame.columns:
            raise ValueError("'raw_content' column is missing!")

        parsed_entries = []
        for index, row in data_frame.iterrows():
            try:
                if pd.notna(row["raw_content"]):
                    content = str(row["raw_content"]).strip()
                    parsed_json = self.try_parse_json(content, index)
                    parsed_entries.append({
                        "original_index": index,
                        "parsed_data": parsed_json
                    })
            except Exception as error:
                print(f"Error in row {index}: {error}")

        if not parsed_entries:
            raise ValueError("No data parsed successfully!")

        print(f"Successfully parsed {len(parsed_entries)} entries.")
        return parsed_entries

    def try_parse_json(self, content: str, index: int) -> Dict:
        """Attempt to parse JSON content, fixing issues if needed."""
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            print(f"JSON error in row {index}, attempting to fix...")
            start_idx = content.find("{")
            if start_idx != -1:
                brace_count = 0
                end_idx = start_idx
                for i in range(start_idx, len(content)):
                    if content[i] == "{":
                        brace_count += 1
                    elif content[i] == "}":
                        brace_count -= 1
                        if brace_count == 0:
                            end_idx = i
                            break
                clean_content = content[start_idx:end_idx + 1]
                return json.loads(clean_content)
            else:
                raise

    def get_or_create_id(self, category: str, value: str) -> int:
        """Get or create an ID for a given category."""
        if not value or pd.isna(value):
            value = "unknown" if category == "comm_type" else "No Subject"

        if value not in self.lookup_tables[category]:
            new_id = self.id_counters[category]
            self.lookup_tables[category][value] = new_id
            self.dimensions[category].append({f'{category}_id': new_id, category: value})
            self.id_counters[category] += 1

        return self.lookup_tables[category][value]

    def extract_emails(self, items: Any) -> List[str]:
        """Extract emails from a list of items."""
        emails = set()
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict) and 'email' in item:
                    if item['email'] and not pd.isna(item['email']):
                        emails.add(str(item['email']).strip().lower())
                elif isinstance(item, str) and '@' in item:
                    emails.add(str(item).strip().lower())
        elif isinstance(items, str) and '@' in items:
            emails.add(str(items).strip().lower())
        return list(emails)

    def transform_data(self, parsed_data: List[Dict], input_file: str) -> None:
        """Transform parsed data into fact and dimension tables."""
        print("Step 2: Transforming data...")

        original_data = pd.read_csv(input_file) if input_file.endswith(".csv") else pd.read_excel(input_file, engine="openpyxl")

        for item in parsed_data:
            data = item['parsed_data']
            title = data.get('title', '')
            duration = data.get('duration', None) 
            calendar_id = data.get('calendar_id', None)
            audio_url = data.get('audio_url', None)
            video_url = data.get('video_url', None)
            transcript_url = data.get('transcript_url', None)
            date_string = data.get('dateString', None)

            comm_type = self.determine_comm_type(audio_url, video_url, title)
            comm_type_id = self.get_or_create_id("comm_type", comm_type)
            subject_id = self.get_or_create_id("subject", title)
            calendar_id_fk = self.get_or_create_id("calendar", calendar_id)
            audio_id = self.get_or_create_id("audio", audio_url)
            video_id = self.get_or_create_id("video", video_url)
            transcript_id = self.get_or_create_id("transcript", transcript_url)

            speakers = self.extract_emails(data.get('speakers', []))
            participants = self.extract_emails(data.get('participants', []))
            attendees = self.extract_attendees(data.get('meeting_attendees', []))

            host_email = self.get_normalized_email(data.get('host_email', None))
            organizer_email = self.get_normalized_email(data.get('organizer_email', None))

            # Collect all unique emails
            all_emails = set(speakers + participants + attendees)
            if host_email:
                all_emails.add(host_email)
            if organizer_email:
                all_emails.add(organizer_email)

            for email in all_emails:
                self.get_or_create_id("user", email)

            self.add_fact_record(
                item['original_index'], data, original_data, date_string,
                comm_type_id, subject_id, calendar_id_fk, audio_id,
                video_id, transcript_id, duration
            )

            for email in all_emails:
                user_id = self.lookup_tables['user'][email]
                self.bridge_records.append({
                    'comm_id': str(uuid.uuid4()),
                    'user_id': user_id,
                    'isAttendee': email in attendees,
                    'isOrganizer': email == organizer_email,
                    'isParticipant': email in participants,
                    'isSpeaker': email in speakers
                })

        print(f"Transformation completed:")
        print(f"  - Total fact records: {len(self.fact_records)}")
        print(f"  - Total bridge records: {len(self.bridge_records)}")
        print(f"  - Unique users: {len(self.dimensions['user'])}")

    def determine_comm_type(self, audio_url: str, video_url: str, title: str) -> str:
        """Identify the type of communication based on provided data."""
        if audio_url or video_url:
            return "meeting"
        elif '@' in title.lower() or 'email' in title.lower():
            return "email"
        elif 'chat' in title.lower():
            return "chat"
        elif 'call' in title.lower():
            return "call"
        return "unknown"

    def get_normalized_email(self, email: Any) -> str:
        """Normalize email to lowercase and stripped format."""
        if email and not pd.isna(email):
            return str(email).strip().lower()
        return None

    def extract_attendees(self, attendees_data: Any) -> List[str]:
        """Extract email addresses from attendees' data."""
        attendees = []
        if isinstance(attendees_data, list):
            for attendee in attendees_data:
                if isinstance(attendee, dict) and 'email' in attendee:
                    email = attendee.get('email')
                    if email and not pd.isna(email):
                        attendees.append(self.get_normalized_email(email))
        return attendees

    def add_fact_record(self, original_index: int, data: Dict, original_data: pd.DataFrame, date_string: str,
                        comm_type_id: int, subject_id: int, calendar_id_fk: int, audio_id: int, video_id: int, 
                        transcript_id: int, duration: Any) -> None:
        """Add a new record to the fact communication table."""
        comm_id = str(uuid.uuid4())
        original_row = original_data.iloc[original_index]

        self.fact_records.append({
            'comm_id': comm_id,
            'raw_id': data.get('id', ''),
            'source_id': original_row.get('source_id', ''),
            'comm_type_id': comm_type_id,
            'subject_id': subject_id,
            'calendar_id': calendar_id_fk,
            'audio_id': audio_id,
            'video_id': video_id,
            'transcript_id': transcript_id,
            'datetime_id': date_string,
            'ingested_at': original_row.get('ingested_at', ''),
            'processed_at': original_row.get('processed_at', ''),
            'is_processed': original_row.get('is_processed', True),
            'raw_title': data.get('title', ''),
            'raw_duration': duration  # Now duration is correctly defined
        })

    def save_to_excel(self, output_file: str) -> None:
        """Export all tables to an Excel file."""
        print("Step 3: Exporting to Excel...")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for category, dimension in self.dimensions.items():
                if dimension:
                    pd.DataFrame(dimension).to_excel(writer, sheet_name=f'dim_{category}', index=False)
                    print(f"  - {category} dimension: {len(dimension)} rows")
            
            if self.fact_records:
                pd.DataFrame(self.fact_records).to_excel(writer, sheet_name='fact_communication', index=False)
                print(f"  - fact communication: {len(self.fact_records)} rows")
            
            if self.bridge_records:
                pd.DataFrame(self.bridge_records).to_excel(writer, sheet_name='bridge_comm_user', index=False)
                print(f"  - bridge communication: {len(self.bridge_records)} rows")

        print(f"Export completed: {output_file}")

    def execute_etl(self, input_file: str, output_file: str) -> None:
        """Execute the complete ETL process."""
        print("=== Starting ETL Process ===")
        
        try:
            parsed_data = self.load_and_parse_data(input_file)
            self.transform_data(parsed_data, input_file)
            self.save_to_excel(output_file)

            print("=== ETL Process Completed ===")
            print(f"Results Summary:")
            print(f"  Dimension tables: {len(self.dimensions)}")
            print(f"  Fact table: {len(self.fact_records)} rows")
            print(f"  Bridge table: {len(self.bridge_records)} rows")
            print(f"  Total unique users: {len(self.dimensions['user'])}")

        except Exception as error:
            print(f"Error during ETL process: {error}")
            raise

if __name__ == "__main__":
    etl_process = CommunicationDataETL()
    
    print(f"Current directory: {os.getcwd()}")

    try:
        if not os.path.exists(input_file_path):
            print(f"File not found: {input_file_path}")
            print("Available files:")
            directory = os.path.dirname(input_file_path)
            if os.path.exists(directory):
                for file in os.listdir(directory):
                    if file.endswith(('.xlsx', '.csv')):
                        print(f"  - {file}")
            else:
                print(f"Directory does not exist: {directory}")
            exit(1)

        etl_process.execute_etl(input_file_path, output_file_path)

    except FileNotFoundError:
        print(f"File not found: {input_file_path}")
        print("Please check the file path.")
    except Exception as error:
        print(f"Error during ETL execution: {error}")
        import traceback
        traceback.print_exc()
