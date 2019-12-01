import optparse as op
import sys
import requests
import json
import xlwt
import time

parser = op.OptionParser()
args = parser.parse_args()
if len(args[1][0]) != 85:
    print('You have to insert ACCESS_TOKEN after calling programm.')
    sys.exit()
ACCESS_TOKEN = args[1][0]


class User:
    """Allows to use modified searching in a VK users`s friends list

    User(id, ACCESS_TOKEN, V)
        id - VK user`s id where need to do searching
        ACCESS_TOKEN - user`s ACCESS_TOKEN via user`s access is got from VK
        V - version of VK API

    >>>Denis = User('122049063', ACCESS_TOKEN, '5.103')
    """
    def __init__(self, id, ACCESS_TOKEN, V):
        self.URL_TEMPLATE = 'https://api.vk.com/method/{METHOD_NAME}?{PARAMETERS}&access_token={ACCESS_TOKEN}&v={V}'
        self.id = id
        self.ACCESS_TOKEN = ACCESS_TOKEN    
        self.V = V

    def getinfo(self):
        """Return info in JSON format about user which ACCESS_TOKEN is in the object.

        >>>Denis = User('122049063', ACCESS_TOKEN, '5.103')
        >>>Denis.getinfo()
        {'response': [{'id': 122049063, 'first_name': 'Денис', 'last_name': 'Мещеряков', 'is_closed': False, 'can_access_closed': True}]}
        """
        api_url = self.URL_TEMPLATE.format(METHOD_NAME='users.get', PARAMETERS='', ACCESS_TOKEN=self.ACCESS_TOKEN, V=self.V)
        response = requests.get(api_url)
        return json.JSONDecoder().decode(response.text)

    def getfriends(self, recursive_level=0):
        """Return user`s VK friends list for which the id was inserted in JSON format.

        >>>Denis = User('122049063', ACCESS_TOKEN, '5.103')
        >>>Denis.getfriends()
        {'response': {'count': 127, 'items': [{'id': 172343456, 'first_name': 'Alex', 'last_name': 'Kosterev', 'is_closed': True,
        'can_access_closed': True, 'domain': 'id172343456', 'online': 0, 'track_code': '24db9be0fe-rD38cPwXDl6tuKwgHf3kC1yr-n8TOQg--JwQlQBAQhKQ4TB0MDZfUwN_H'},...]}}
        """
        api_url = self.URL_TEMPLATE.format(METHOD_NAME='friends.get', ACCESS_TOKEN=self.ACCESS_TOKEN, V=self.V,
                                           PARAMETERS='user_id={0}&order={1}&fields={2}'.format(self.id, 'name', 'domain',))
        response = requests.get(api_url)
        time.sleep(0.15)
        response = json.JSONDecoder().decode(response.text)
        return self.__get_recursive(response, recursive_level)

    def getfriends_bynames(self, *args, recursive_level=0):
        """Return result of searching by names in user`s VK friends list for which the id was inserted in JSON format.

        >>>Denis = User('122049063', ACCESS_TOKEN, '5.103')
        >>>Denis.getfriends_bynames('Андрей', 'Екатерина')
        {'response': {'count': 5, 'items': [{'id': 2205229, 'first_name': 'Андрей', 'last_name': 'Barmaley', 'is_closed': False,
        'can_access_closed': True, 'domain': 'andreybarmaley7', 'online': 0, 'track_code': 'bc097c04sgmkK07HdkriAFA_KJvDmU0rQPdiWYFVvsbtT2ykgb3fYqlKecQhfexRArqTPw'},...]}}
        """
        if len(args) == 1:
            api_url = self.URL_TEMPLATE.format(METHOD_NAME='friends.search', ACCESS_TOKEN=self.ACCESS_TOKEN, V=self.V,
                                               PARAMETERS='user_id={0}&q={1}'.format(self.id, args[0]))
            response = requests.get(api_url)
            return json.JSONDecoder().decode(response.text)
        return self.__getuserslist(*args, recursive_level = recursive_level)

    def getfriends_exceptnames(self, *args, recursive_level=0):
        """Return result of searching except names in user`s VK friends list for which the id was inserted in JSON format.

        >>>Denis = User('122049063', ACCESS_TOKEN, '5.103')
        >>>Denis.getfriends_exceptnames('Андрей', 'Екатерина')
        ...
        """
        return self.__getuserslist(*args, _except = True, recursive_level = recursive_level)

    def __getuserslist(self, *args, _except=False, recursive_level=0):
        itemstoremove = []
        friendsdict = self.getfriends(recursive_level = recursive_level)
        friendsdict['response']['items'] = list(list(list(item for item in person.items())) for person in friendsdict['response']['items'])
        for item in friendsdict['response']['items']:
            if _except == False:
                if item[1][1] not in args:
                    itemstoremove.append(item)
            else:
                if item[1][1] in args:
                    itemstoremove.append(item)
        for item in itemstoremove:
            friendsdict['response']['items'].remove(item)
        friendsdict['response']['items'] = list(dict(person) for person in friendsdict['response']['items'])
        friendsdict['response']['count'] = len(friendsdict['response']['items'])
        return friendsdict

    def __get_recursive(self, response, recursive_level=0):
        assert recursive_level >= 0, 'Recursive_level must be non-negative or zero.'
        if recursive_level == 0:
            return response
        for num, user in enumerate(response['response']['items']):
            exec('user{num} = User(user["id"], self.ACCESS_TOKEN, self.V)\n'
                'response{num} = user{num}.getfriends(recursive_level = recursive_level - 1)\n'
                'response["response"]["items"].extend(response{num}["response"]["items"])\n'
                'response["response"]["count"] += response{num}["response"]["count"]'.format(num = num))
        return response

def main():
    print(User('122049063', ACCESS_TOKEN, '5.103').getfriends_bynames('Андрей', 'Екатерина'))
    users_json = User('122049063', ACCESS_TOKEN, '5.103').getfriends_bynames('Андрей', 'Екатерина')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Users')
    ws.write(0, 0, 'FIRSTNAME')
    ws.write(0, 1, 'LASTNAME')
    ws.write(0, 2, 'ID')
    for i, user in enumerate(users_json['response']['items']):
        ws.write(i + 1, 0, user['first_name'])
        ws.write(i + 1, 1, user['last_name'])
        ws.write(i + 1, 2, 'https://vk.com/id{0}'.format(user['id']))
    wb.save('.\\Users.xls')


    
main()
