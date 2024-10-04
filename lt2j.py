#!/usr/bin/python3

import sys
import os.path
import pyexcel as pyex
import argparse
import json
from jira import JIRA
from tzlocal import get_localzone
from datetime import *

# Built-in configuration {{{
#
# The field mapping
fmap = {
  "date": 0,
  "time_from": 1,
  "time_to": 2,
  "duration": 3,
  "issue_id": 4,
  "description": 5,
}

# }}}
# Parse the command-line arguments {{{
#
# Create a parser, add all the available arguments to it
parser = argparse.ArgumentParser(prog="lt2j", description="Read a spreadsheet containing worklog entries, and add them to Jira.")
parser.add_argument("-d", "--debug", action="store_true", help="Run in debug mode.")
parser.add_argument("-y", "--yolo", "--yes", action="store_true", help="Yes, really do create worklog entries in jira.")
parser.add_argument("-f", "--file", required=True, help="Path to the file with the spreadsheet")
parser.add_argument("-n", "--sheet", required=True, help="Name of the sheet (tab) where worklog entries are stored.")
parser.add_argument("-s", "--start", type=int, default=0, help="Start importing frow this row number.")
parser.add_argument("-e", "--end", type=int, default=0, help="End importing at this row number.")
parser.add_argument("-u", "--jira-url", required=True, help="The URL where Jira is hosted at.",)
parser.add_argument("-t", "--jira-token", required=True, help="The private token for Jira")
parser.add_argument("command", choices=["create", "remove"], default="create", help="Create or remove worklogs from Jira")
args = parser.parse_args()

# }}}
# Do some sanity checking{{{
#
# Check for common errors in the configuration that we've got
if not args.yolo:
  print(f"The yolo mode is *NOT* on! Not doing anything, just reporting what would have been done.")

# Does the file exist?
if not os.path.isfile(args.file):
  print("ERROR: Spreadsheet file '" + args.file + "' does not exist, aborting.")
  exit(1)

# Check if row indexes are valid
if args.end > 0 and args.end < args.start:
  print("ERROR: Row indexes are messed up, aborting")
  exit(1)
if args.end == args.start:
  print("ERROR: Can't start and end at the same row index, aborting.")
  exit(1)

# }}}
# Main stuff {{{
#
# Load data from file
sheet = pyex.get_sheet(
  file_name = args.file,
  sheet_name = args.sheet,
  start_row = args.start,
  row_limit = args.end - args.start,
  skip_empty_rows = True,
)
if args.debug:
  print(f"DEBUG: Loaded data from file {args.file}, from sheet {args.sheet}, starting at row {args.start}. Read {sheet.number_of_rows()} rows with {sheet.number_of_columns()} columns from sheet", file=sys.stderr)

# Authenticate to jira
jira = JIRA(server = args.jira_url, token_auth = args.jira_token)
jira_user_key = jira.myself()["key"]
if args.debug: print(f"DEBUG: Authenticated as: jira_user_key={jira_user_key}, {json.dumps(jira.myself(), indent=2, sort_keys=True)}", file=sys.stderr)

# Add worklogs
if args.command == "create":
  for row in sheet:
    # Check for sanity, and extract relevant data from spreadsheet into
    # variables.
    if args.debug: print(f"DEBUG: Processing row: {row}", file=sys.stderr)
    if len(row) < len(fmap):
      print(f"Warning: Row {row} has {len(row)} fields, but we require at least {len(fmap)}, skipping this row.")
      continue
    date = row[fmap["date"]]
    time_from = row[fmap["time_from"]]
    time_to = row[fmap["time_to"]]
    duration = row[fmap["duration"]]
    issue_id = row[fmap["issue_id"]]
    description = row[fmap["description"]]

    # Calculate dates and times, do some debug output
    if args.debug: print(f"DEBUG: parsed from row: date={date}, time_from={time_from}, time_to={time_to}, duration={duration}, issue_id={issue_id}, description={description}", file=sys.stderr)
    started_dt = datetime(year = date.year, month = date.month, day = date.day, hour = time_from.hour, minute = time_from.minute, second = time_from.second, tzinfo = get_localzone())
    duration_sec = duration.second + duration.minute * 60 + duration.hour * 3600
    if args.debug: print(f"DEBUG: calculated dates and durations: started_dt={started_dt}, duration_sec={duration_sec}", file=sys.stderr)

    # Add worklogs to jira (or not, depending on mode)
    if args.yolo:
      print(f"Adding worklog (started_dt={started_dt}, duration_sec=@{duration_sec}, description={description[:36]!r}) for issue (issue_id={issue_id})")
      jira.add_worklog(issue = issue_id, started = started_dt, timeSpentSeconds = duration_sec, comment = description)
    else:
      print(f"Would add worklog (started_dt={started_dt}, duration_sec=@{duration_sec}, description={description[:24]!r}) for issue (issue_id={issue_id})")

# Remove worklogs
elif args.command == "remove":
  # print(f"The '{args.command}' command is not yet implemented, aborting.")
  # exit(1)

  # Set up dicts that will be used to cache data for issues and worklogs, then
  # cycle through all the rows in the sheet.
  worklogs = {}
  issues = {}
  for row in sheet:
    # Check for sanity, and extract relevant data from spreadsheet into
    # variables.
    if args.debug: print(f"DEBUG: Processing row: {row}", file=sys.stderr)
    if len(row) < len(fmap):
      print(f"Warning: Row {row} has {len(row)} fields, but we require at least {len(fmap)}, skipping this row.")
      continue
    date = row[fmap["date"]]
    time_from = row[fmap["time_from"]]
    time_to = row[fmap["time_to"]]
    duration = row[fmap["duration"]]
    issue_id = row[fmap["issue_id"]]
    description = row[fmap["description"]]

    # Calculate dates and times
    started_dt = datetime(year = date.year, month = date.month, day = date.day, hour = time_from.hour, minute = time_from.minute, second = time_from.second, tzinfo = get_localzone())
    duration_sec = duration.second + duration.minute * 60 + duration.hour * 3600
    if args.debug: print(f"DEBUG: calculated dates and durations: started_dt={started_dt}, duration_sec={duration_sec}", file=sys.stderr)

    # Have we seen this issue yet? If not, add it to "issues" dict, retrieve a
    # list of all worklogs that are associated to it, then retrieve data for
    # all these worklogs into the "worklogs" dict
    if not issue_id in issues:
      issues[issue_id] = jira.worklogs(issue_id)
      print(f"Loading worklogs for issue (issue_id={issue_id}): ", end="", flush=True)
      if args.debug: print(f"DEBUG: These are all worklogs for issue_id={issue_id}: {issues[issue_id]}", file=sys.stderr)

      # Retrieve all worklogs for this issue, store into "worklogs" dict
      for worklog_id in issues[issue_id]:
        print(f".", end="", flush=True)
        worklogs[worklog_id] = jira.worklog(issue=issue_id, id=worklog_id)
        if args.debug: print(f"DEBUG: This is worklog {worklog_id} for issue {issue_id}: {json.dumps(worklogs[worklog_id].raw, indent=2, sort_keys=True)}", file=sys.stderr)
      print()

    # Figure out what is the correct worklog_id for the worklog that this row
    # describes.
    for worklog in worklogs:
      worklog_author_key = worklog.raw["author"]["key"]
      worklog_started_dt = datetime.strptime(worklog.raw["started"], "%Y-%m-%dT%H:%M:%S.000%z")
      worklog_duration_sec = int(worklog.raw["timeSpentSeconds"])
      worklog_id = worklog.raw["id"]
      if args.debug: print(f"DEBUG: checking if worklog (worklog_id={worklog_id}, worklog_author_key={worklog_author_key}, worklog_started_dt={worklog_started_dt}, worklog_duration_sec={worklog_duration_sec}) matches with worklog (started_dt={started_dt}, duration_sec={duration_sec}). ", file=sys.stderr, end="")
      if worklog_author_key == jira_user_key and worklog_started_dt == started_dt and worklog_duration_sec == duration_sec:
        if args.debug: print(f"YES!", file=sys.stderr)
        if args.yolo:
          print(f"Deleting worklog (worklog_id={worklog_id}) for issue (issue_id={issue_id})")
          worklog.delete()
        else:
          print(f"Would delete worklog (worklog_id={worklog_id}) for issue (issue_id={issue_id})", file=sys.stderr)
        continue
      else:
        if args.debug: print(f"No.", file=sys.stderr)

# Unknown action
else:
  print(f"Unknown command {args.command}, aborting.")
  exit(1)

# }}}

# vim: set ts=2 sw=2 sts=2 et cc=80 fdl=0 fdm=marker:
