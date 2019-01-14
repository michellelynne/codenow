import argparse
import csv
import json
import re
import textwrap
from datetime import datetime

from cached_property import cached_property
from pyexcel_ods import get_data


class User:

    def __init__(self, email, department_id=None, historical_matches=None):
        self.email = email
        self.department_id = department_id
        self.historical_matches = historical_matches if historical_matches else []

    def get_possible_matches(self, available_emails, user_departments):
        """Get a list of matches for this user.
        Match must be signed up for this month.
        Match must not be in same department.
        Match must not be matched before.

        Args:
            available_emails(List<str>): List of people who signed up for matches currently.
            user_departments(dict): Dict of users and their department ids.

        Returns:
            List<str>: List of matches for this user.

        """
        _possible_matches = []
        previous_matches = [self.email] + self.historical_matches
        for possible_match in (set(available_emails) - set(previous_matches)):
            if user_departments.get(possible_match) != self.department_id:
                _possible_matches.append(possible_match)
        return sorted(_possible_matches)


class Matcher:

    def __init__(self, requests_filename, employees_filename, matches_filename):
        """Match users to other other available users.

        Args:
            matches_filename(str): Filename of matches.
            employees_filename(str): Filename of employee objects.
            requests_filename(str): Filename of employee objects.

        """
        self.requests_filename = requests_filename
        self.employees_filename = employees_filename
        self.matches_filename = matches_filename

    @cached_property
    def historical_matches(self):
        """Get a dictionary of users and their previous matches."""
        _historical_matches = {}
        email_search = re.compile('([\w\.\_]*@chownow.com)')
        matches_dict = get_data(self.matches_filename)
        for tab, data in matches_dict.items():
            email_locations = []
            for loc, item in enumerate(data[1]):
                if email_search.match(str(item)):
                    email_locations.append(loc)
            for row in data[1:]:
                if not row:
                    continue
                email_key1 = row[email_locations[0]].lower()
                email_key2 = row[email_locations[1]].lower()
                _historical_matches.setdefault(email_key1, [])
                _historical_matches[email_key1].append(email_key2)
        return _historical_matches

    @cached_property
    def user_departments(self):
        """Get a dictionary of users and their departments."""
        _user_departments = {}
        with open(self.employees_filename, 'r') as employee_file:
            employees = json.load(employee_file)
        for employee in employees['objects']:
            email = employee['email'].lower()
            if email:
                _user_departments[email] = employee['department_id']
        return _user_departments

    @cached_property
    def available_emails(self):
        """Get a list of email addresses who want matches from the requests file."""
        _available_emails = []
        with open(self.requests_filename, 'r') as csvfile:
            requests_reader = csv.DictReader(csvfile)
            for row in requests_reader:
                match_me = row['Match me up for a chat!'].lower()
                if 'yes' not in match_me:
                    continue
                email = row['Email Address']
                _available_emails.append(email)
        return sorted(_available_emails)

    @cached_property
    def users(self):
        """Get a list of User models."""
        _users = []
        for email in self.available_emails:
            _user_historical_matches = self.historical_matches.get(email, [])
            _user_department_id = self.user_departments.get(email, None)
            user = User(email=email,
                        department_id=_user_department_id,
                        historical_matches=_user_historical_matches)
            _users.append(user)
        return _users

    @cached_property
    def matches(self):
        """Get a list of matches."""
        _matches = []
        available_emails = set(self.available_emails)
        already_matched = []
        for user in self.users:
            if user.email in set(already_matched):
                continue
            user_matches = user.get_possible_matches(available_emails, self.user_departments)
            last_match = user_matches[0]
            already_matched.append(last_match)
            available_emails -= set([last_match, user.email])
            _matches.append({
                'Email Address': user.email,
                'Meet your Match': last_match,
            })
        return _matches

    def export_matches(self):
        """Export matches as a csv file."""
        csv_file_name = 'CoffeeNow_{}.csv'.format(datetime.now().strftime('%Y%m%d'))
        with open(csv_file_name, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=self.matches[0].keys())
            writer.writeheader()
            for match in self.matches:
                writer.writerow(match)


def get_args():
    description = textwrap.dedent('''
    CoffeeNow aims to bring people together who haven't met.
    Given a history of matches and departments, matches two people together/

    Examples: 

    coffee_now.py -r data/responses.csv -m data/historical_matches.ods -e data/employee_objects.json
    ''')

    parser = argparse.ArgumentParser(
        description=description,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('-r', '--requests_filename',
                        help='Users who are requesting a match as CSV file.', required=True)
    parser.add_argument('-m', '--matches_filename',
                        help='Historical matches of users as ODS file.', required=True)
    parser.add_argument('-e', '--employees_filename',
                        help='Employee objects file from Zenefits as JSON.', required=True)
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    matcher = Matcher(args.requests_filename, args.employees_filename, args.matches_filename)
    matcher.export_matches()