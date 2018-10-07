#!/usr/bin/env python3
import configparser
import csv
import email.message
import getpass
import os
import pickle
import smtplib
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path


os.system('')  # enable ANSI escape code on Windows *facepalm*

# sorry for the global variable... at least it make sense for a config
config = configparser.ConfigParser(
    converters={'datetime': lambda s: datetime.strptime(s, '%Y-%m-%d')}
)
config.read('config.ini')
config_m = config['MeetNGo']
#config_p = config['Parainage']
config_t = config['Tandem']



def ask(question, default=True):
    """Ask yes-no question, fallback to yes if `default` is set."""
    question += ' [O/n] ' if default else ' [o/N] '
    ans = input(question).strip().lower()
    if ans == 'o':
        return True
    elif ans == 'n':
        return False
    else:
        return default


def compat(s):
    """Return the string in lowercase and without accents."""
    return ''.join(c for c in unicodedata.normalize('NFKD', s)
                   if unicodedata.category(c) != 'Mn')\
             .casefold().replace(' ', '').replace('-', '')


def find_file(name, folder='.'):
    """Find a file that contain `name` in its filename."""
    files = list(Path(folder).glob('*{}*'.format(name)))

    if len(files) == 1:
        return files[0]
    elif len(files) < 1:
        input(color('\nAucun fichier ne correspond à « {} ».'.format(name), 'red'))
        exit()

    print('\nIl y a {} fichiers qui correspondent à « {} » :'.format(len(files), name))
    print('\n'.join(['  {}) {}'.format(i+1, f) for i, f in enumerate(files)]))
    try:
        ans = int(input('\nSaisir votre choix (défaut=1): ')) - 1
    except (ValueError):
        ans = 0
    print()
    return files[ans] if 0 <= ans and ans < len(files) else files[0]
    return files[ans] if 0 <= ans and ans < len(files) else files[0]


def generate_emails(data, subject, template_path):
    """Generate email from a predefined template."""
    with open(template_path) as tmpl:
        template = tmpl.read()
    emails = []
    for datum in data:
        msg = email.message.EmailMessage()
        msg['To'] = datum['recipient'].email
        msg['Subject'] = subject
        msg.set_content(template.format_map(datum))
        emails.append(msg)
    return emails


def send_emails(emails):
    """Send a list of emails with the provided account."""
    with smtplib.SMTP('smtp.phpnet.org', 587) as smtp:
        smtp.starttls()

        print('\nConnexion à integre-grenoble.org')
        username = input("Nom d'utilisateur : ").strip()
        if not username.endswith('@integre-grenoble.org'):
            username += '@integre-grenoble.org'
        smtp.login(username, getpass.getpass('Mot de passe : '))

        for msg in emails:
            msg['From'] = username
            smtp.send_message(msg)
            print('email envoyé à {}'.format(msg['To']))


def color(string=None, fg=None, bg=None, style=None, *args):
    """Color string with ANSI escape codes."""
    colors = ['black', 'red', 'green', 'yellow',
              'blue', 'magenta', 'cyan', 'white']
    styles = {'reset': 0, 'bold': 1, 'bright': 1, 'underline': 4, 'reverse': 7}

    param = []
    for attr in [style] + list(args):
        if attr in styles:
            param.append(str(styles[attr]))
    if fg in colors:
        param.append(str(30 + colors.index(fg)))
    if bg in colors:
        param.append(str(40 + colors.index(bg)))

    srg = '\x1B[{}m'.format(';'.join(param))
    return srg if string is None else srg + string + '\x1B[m'


class Menu:
    def __init__(self):
        self.title = ''
        self.title_color = 'blue'
        self.subtitle = ''
        self.subtitle_color = 'magenta'
        self.message = ''
        self.items = []
        self.prompt = '>> '

    def display(self, only_once=False):
        """Show a CLI menu parameterized by the class attributes."""
        while True:
            print('\x1Bc')  # reset display
            print(color(self.title, self.title_color, '', 'bold'))
            print(color('{:>80}'.format(self.subtitle), self.subtitle_color))
            if self.message:
                print(self.message)

            print()  # give your menu a bit of space
            for i, item in enumerate(self.items):
                print('{:>2} - {}'.format(i + 1, item[0]))
            choice = input(self.prompt)

            try:
                if int(choice) < 1:
                    raise ValueError
                self.items[int(choice) - 1][1]()
            except (ValueError, IndexError):
                pass

            if only_once:
                break


class Group(set):

    def __init__(self, person_class, *args):
        self.person_class = person_class
        super().__init__(*args)

    def append(self, person):
        """Add a person to the group, ask user if already into."""
        exist = False
        for other in self:
            if not exist and person.look_like(other):
                print('\n{}\n{}'.format(other, person))
                if ask('Ces personnes sont elles la même personne ?'):
                    print('Copie suprimée !')
                    exist = True
                else:
                    print('Réponse prise en compte. Ces deux personnes sont différentes.')
        if not exist:
            self.add(person)

    def load(self, filename, nb_title_row=1, ignore_before=datetime.min):
        """Load a csv file into the group."""
        with open(filename) as f:
            reader = csv.reader(f)
            for _ in range(nb_title_row):
                next(reader)  # discard title row(s)
            for row in reader:
                if datetime.strptime(row[0][:10], '%Y/%m/%d') > ignore_before:
                    self.append(self.person_class(row))



####################          Meet'N'Go             ####################

class Mentor:

    def __init__(self, row):
        """Initialize a mentor from a csv row."""
        self.surname = row[1].strip()
        self.name = row[2].strip()
        self.email = row[3].strip()
        self.lang = set(row[7].split(';'))  # TODO: strip each lang
        self.country = row[8].strip()
        self.city = row[9].strip()
        self.university = row[10].strip()
        self.present_the_23 = True if row[16] == 'Oui / Yes' else False
        self.helping_the_23 = True if row[17] == 'Oui / Yes' else False

        self.mentees = []

    @property
    def languages(self):
        return ', '.join(self.lang)

    def look_like(self, other):
        """Return `True` if `other` might be the same person."""
        return (compat(self.email) == compat(other.email)
                or (compat(self.name) == compat(other.name)
                    and compat(self.surname) == compat(other.surname)))

    # TODO: understand and change email generation
    def generate_email(self):
        """Generate emails form a predefined template."""
        emails = ''
        template_path = config['Templates']['folder'] + '/'
        template_path += config['Templates']['mentors']
        with open(template_path) as tmpl:
                template = tmpl.read()
        for mentee in self.mentees:
            emails += template.format(recipient=self, mentee=mentee)
        return emails

    # TODO: must save on a csv-like format
    def save(self):
        """Pickle mentor information for later use."""
        path = Path(config['Data']['top folder'])
        path = path / config['Data']['mentors folder']
        if not path.exists():
            path.mkdir(parents=True)
        path = path / '{}.{}.pickle'.format(compat(self.name),
                                            compat(self.surname))
        with path.open('wb') as f:
            pickle.dump(self, f, protocol=pickle.HIGHEST_PROTOCOL)

    def __str__(self):
        return ' - {name} {surname}, {email}, a étudier en {country}, parle {lang}.'.format_map(self.__dict__)


class Mentee:

    print(email)
    def __init__(self, row):
        """Initialize a mentee from a csv row."""
        self.surname = row[1].strip()
        self.name = row[2].strip()
        self.email = row[3].strip()
        self.lang = set(row[7].split(';'))  # TODO: strip each lang
        self.country = row[8].strip()
        self.city = row[9].strip()
        self.university = row[10].strip()
        self.present_the_23 = True if row[16] == 'Oui / Yes' else False

        self.mentor = None

    @property
    def languages(self):
        return ', '.join(self.lang)

    def look_like(self, other):
        """Return `True` if `other` might be the same person."""
        return (compat(self.email) == compat(other.email)
                or (compat(self.name) == compat(other.name)
                    and compat(self.surname) == compat(other.surname)))

    def find_mentor(self, mentors_group):
        """Find the most suitable mentors among `mentors_group`."""
        # only keep mentor that have been the the desired country
        mentors = [m for m in mentors_group
                   if compat(self.country) == compat(m.country)]
        if len(mentors) > 0:
            # if some mentors remain, try to limit to the desired city
            by_city = [m for m in mentors
                       if compat(self.city) == compat(m.city)]
            if len(by_city) > 0:
                # if some mentors remain, try to limit to the desired univ
                mentors = by_city
                by_univ = [m for m in mentors
                           if compat(self.university) == compat(m.university)]
                if len(by_univ) > 0:
                    mentors = by_univ
            # among the remaining mentors, choose the one with less mentees
            mentors.sort(key=lambda mentor: len(mentor.mentees))
            self.mentor = mentors[0]
            # don't forget to add self to its mentor's mentees
            self.mentor.mentees.append(self)

    # TODO: understand and change email generation
    def generate_email(self):
        """Generate email form predefined templates."""
        template_path = config['Templates']['folder'] + '/'
        if self.mentor is None:
            template_path += config['Templates']['alone mentees']
            with open(template_path) as tmpl:
                email = tmpl.read().format(recipient=self)
        else:
            template_path += config['Templates']['mentees']
            with open(template_path) as tmpl:
                email = tmpl.read().format(recipient=self, mentor=self.mentor)
        return email

    # TODO: must save on a csv-like format
    def save(self):
        """Pickle mentee information for later use."""
        path = Path(config['Data']['top folder'])
        path = path / config['Data']['mentees folder']
        if not path.exists():
            path.mkdir(parents=True)
        path = path / '{}.{}.pickle'.format(compat(self.name),
                                            compat(self.surname))
        with path.open('wb') as f:
            pickle.dump(self, f, protocol=pickle.HIGHEST_PROTOCOL)

    def __str__(self):
        return ' - {name} {surname}, {email}, veut aller en {country}, parle {lang}.'.format_map(self.__dict__)


def do_meetngo():
    print('\x1Bc')
    print(color('''\
   _____                 __ /\_______ /\ ________
  /     \   ____   _____/  |)/\      \)//  _____/  ____
 /  \ /  \_/ __ \_/ __ \   __\/   |   \/   \  ___ /  _ \\
/    Y    \  ___/\  ___/|  | /    |    \    \_\  (  <_> )
\____|__  /\___  >\___  >__| \____|__  /\______  /\____/
        \/     \/     \/             \/        \/\n''', 'blue', '', 'bold'))

    print(config_m['répertoire modèles mails'])

    # TODO: last run in the config

    # create a mentor group and fill it with old data (if the user agree)
    mentors = Group(Mentor)
    # TODO: old data
    # add new mentors from csv
    mentors_file = find_file(config_m['fichier parrains'],
                             config_m['répertoire csv'])
    print('« {} » sera utilisé ajouter de nouveaux parrains.'.format(mentors_file))
    mentors.load(mentors_file, last_run)

    # create a mentor group and fill it with old data (if the user agree)
    mentees = Group(Mentees)
    # TODO: old data
    # add new mentors from csv
    mentees_file = find_file(config_m['fichier filleuls'],
                             config_m['répertoire csv'])
    print('« {} » sera utilisé ajouter de nouveaux filleuls.'.format(mentees_file))
    mentees.load(mentees_file, last_run)

    # try to find a mentor for each mentee
    for mentee in mentees:
        mentee.find_mentor(mentors)



####################          Parainage             ####################

#def do_parainage():
#    print('\x1Bc')
#    print(color('''\
#__________                     .__
#\______   \_____ ____________  |__| ____ _____     ____   ____
# |     ___/\__  \\_  __ \__  \ |  |/    \\__  \   / ___\_/ __ \\
# |    |     / __ \|  | \// __ \|  |   |  \/ __ \_/ /_/  >  ___/
# |____|    (____  /__|  (____  /__|___|  (____  /\___  / \___  >
#                \/           \/        \/     \//_____/      \/\n''', 'blue', '', 'bold'))

#    pass



####################          Tandem                ####################

class Student:

    title_row = ["Timestamp", "Username", "Prénom", "Nom de famille", "Votre Nationalité ?", "Langues parlées", "Langues que tu veux apprendre", "Age", "Université / University", "A partir de quand êtes vous disponible ?"]

    def __init__(self, row):
        """Initialize a student from a csv row."""
        self.email = row[1].strip()
        self.first_name = row[2].strip()
        self.family_name = row[3].strip()
        self.nationality = row[4]
        self.known_lang = row[5].split(';')
        self.wanted_lang = row[6].split(';')
        self.age = int(row[7].strip())
        self.university = row[8]
        self.avail = row[9]

        self.alreay_alone = False
        self.partner = None

    @property
    def known_languages(self):
        """Return a printable list of known language."""
        return ', '.join(self.known_lang)

    @property
    def wanted_languages(self):
        """Return a printable list of wanted language."""
        return ', '.join(self.wanted_lang)

    @property
    def as_row(self):
        """Return a students as an row that can be writen in a   csv file."""
        return [date.today().strftime('%Y/%m/%d'), self.family_name,
                self.first_name, self.email, self.nationality,
                ';'.join(self.known_lang), ';'.join(self.wanted_lang), self.age,
                self.gender, self.university, self.avail]

    def __str__(self):
        return ' - {first_name} {family_name}, {email}, parle {known_lang} et veut apprendre {wanted_lang}'.format_map(self.__dict__)

    def look_like(self, other):
        """Return `True` if `other` might be the same person."""
        return (compat(self.email) == compat(other.email)
                or (compat(self.first_name) == compat(other.first_name)
                    and compat(self.family_name) == compat(other.family_name)))

    def possible_tandems(self, others):
        tandems = []
        for other in others:
            if id(other) != id(self)\
               and other.partner is None\
               and set(other.known_lang).intersection(self.wanted_lang)\
               and set(other.wanted_lang).intersection(self.known_lang):
                tandems.append(other)
        return sorted(tandems, key=lambda x: (abs(self.age - x.age)))


def do_tandem():
    print('\x1Bc')
    print(color('''\
___________                  .___
\__    ___/____    ____    __| _/____   _____
  |    |  \__  \  /    \  / __ |/ __ \ /     \\
  |    |   / __ \|   |  \/ /_/ \  ___/|  Y Y  \\
  |____|  (____  /___|  /\____ |\___  >__|_|  /
               \/     \/      \/    \/      \/\n''', 'blue', '', 'bold'))

    students = Group(Student)

    if len(list(Path(config_t['répertoire csv']).glob('*{}*'.format(config_t['étudiants seuls'])))) > 0:
        if ask("Un fichier d'étudiants seuls existe, voulez-vous l'utilisez ?"):
            alones_file = find_file(config_t['étudiants seuls'],
                                    config_t['répertoire csv'])
            students.load(alones_file, 1)
    for s in students:
        s.alreay_alone = True

    students_file = find_file(config_t['fichier étudiants'],
                              config_t['répertoire csv'])
    print('« {} » sera utilisé ajouter de nouveaux étudiants.'.format(students_file))
    students.load(students_file, config_t.getint('lignes ignorées', 1))

    tandems = []
    alones = []
    for s in sorted(students, key=lambda x: len(x.possible_tandems(students))*100 + x.age):
        if s.partner is None:
            if len(s.possible_tandems(students)) < 1:
                alones.append(s)
            else:
                s.partner = s.possible_tandems(students)[0]
                s.partner.partner = s
                tandems.append((s, s.partner))
    print('\nIl y a {} tandems, et {} étudiants se retrouvent seuls.'.format(len(tandems), len(alones)))

    if len(list(Path(config_t['répertoire csv']).glob('*{}*'.format(config_t['liste des tandems'])))) > 0:
        input(color("Le fichier {} est sur le point d'être écrasé.".format(config_t['répertoire csv'][2:] + '/' + config_t['liste des tandems']), 'red'))
    with open(config_t['répertoire csv'] + '/' + config_t['liste des tandems'],
              'w', newline='') as tandems_file:
        writer = csv.writer(tandems_file, quoting=csv.QUOTE_ALL)
        writer.writerow(Student.title_row)
        print('Écriture de la liste des tandems dans « {} ».'.format(tandems_file.name[2:]))
        for tandem in tandems:
            writer.writerow([])
            writer.writerow(tandem[0].as_row)
            writer.writerow(tandem[1].as_row)

    with open(config_t['répertoire csv'] + '/' + config_t['étudiants seuls'],
              'w', newline='') as alones_file:
        writer = csv.writer(alones_file, quoting=csv.QUOTE_ALL)
        writer.writerow(Student.title_row)
        print('Écriture de la liste des étudiant seuls dans « {} ».'.format(alones_file.name[2:]))
        for alone in alones:
            writer.writerow(alone.as_row)

    with_partner = []
    for pair in tandems:
        with_partner.append(pair[0])
        with_partner.append(pair[1])

    new_alones = []
    already_alones = []
    for alone in alones:
        if alone.alreay_alone:
            already_alones.append(pair[0])
        else:
            new_alones.append(pair[1])

    emails = []
    if (len(with_partner) > 0 and ask('Voulez-vous envoyer un e-mail aux étudiants qui ont un tandem ?')):
        emails += generate_emails(
            [{'recipient': s, 'tandem': s.partner} for s in with_partner],
            "IntEGre - programme Tandem",  # TODO: mail subject
            config_t['répertoire modèles mails'] + '/' + config_t['modèle mail tandem']
        )
    if (len(new_alones) > 0 and ask('Voulez-vous envoyer un e-mail aux nouveaux étudiants seuls ?')):
        emails += generate_emails(
            [{'recipient': s} for s in new_alones],
            "IntEGre - programme Tandem",  # TODO: mail subject
            config_t['répertoire modèles mails'] + '/' + config_t['modèle mail seul']
        )
    if (len(already_alones) > 0 and ask('Voulez-vous envoyer un e-mail aux étudiants qui était déjà seuls ?')):
        emails += generate_emails(
            [{'recipient': s} for s in already_alones],
            "IntEGre - programme Tandem",  # TODO: mail subject
            config_t['répertoire modèles mails'] + '/' + config_t['modèle mail seul']
        )
    if len(emails) > 0:
        send_emails(emails)



if __name__ == '__main__':
    if len(sys.argv) > 1:
        if 'meetngo' in sys.argv[1].lower():
            do_meetngo()
        elif 'parainage' in sys.argv[1].lower():
            do_parainage()
        elif 'tandem' in sys.argv[1].lower():
            do_tandem()
        else:
            print('Manuel bla bla bla')  # TODO
    else:
        menu = Menu()
        menu.title = '''\
.___        __ ___________ ________
|   | _____/  |\_   _____//  _____/______   ____
|   |/    \   __\    __)_/   \  __\_  __ \_/ __ \\
|   |   |  \  | |        \    \_\  \  | \/\  ___/
|___|___|  /__|/_______  /\______  /__|    \___  >
         \/            \/        \/            \/'''
        menu.subtitle = "Outil Pour les Programmes Annuels d'IntEGre"
        menu.items = [("Meet'N'Go", do_meetngo),
                      #('Parainage', do_parainage),
                      ('Tandem', do_tandem),
                      ('Quitter', exit)]
        do_tandem()#menu.display(only_once=True)

    print()  # final new line, IMHO it look cleaner with it
