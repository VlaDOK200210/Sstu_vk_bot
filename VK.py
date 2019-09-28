import vk_api
import json
import random
import time
import datetime
from openpyxl import load_workbook


def get_button(label, color, payload=""):
    return {
        "action": {
            "type": "text",
            "payload": {"command":"start"},
            "label": label
        },
        "color": color
    }



token_file = open('../token.txt')
token_open = token_file.read
now_week = int(datetime.datetime.now().strftime("%V")) - 35
group_id = "186837700"
token = str(token_open())  # токен сюды
vk = vk_api.VkApi(token=token)
vk._auth_token()
ivcht11 = load_workbook('./rasp/ivcht11.xlsx')  # ПУТЬ К ТАБЛИЦЕ ЭКСЕЛЬ С РАСПИСАНИЕМ
ivcht11_sh = ivcht11["Лист1"]
keyboard = {
    "one_time": False,
    "buttons": [[
        get_button(label="Сегодня", color="positive"),
        get_button(label="Завтра", color="positive")

    ]
    ]
}


keyboard = json.dumps(keyboard, ensure_ascii=False).encode('utf-8')
keyboard = str(keyboard.decode('utf-8'))
#############################################################
p1 = "Первая пара\n" + "8:30 - 9:30\n"
p2 = "Вторая пара\n" + "9:45 - 11:15\n"
p3 = "Третья пара\n" + "11:30 - 13:00\n"
p4 = "Четвертая пара\n" + "14:40 - 15:10\n"
pp = "\n-----------\n"
# Номера пар: r-расписание 1 - четная 2 - нечетная
############################################################################################################
ivcht_r1_1 = [ivcht11_sh['B6'].value, ivcht11_sh['B7'].value, ivcht11_sh['B8'].value, ivcht11_sh['B9'].value,
              ivcht11_sh['B10'].value]
ivcht_r1_2 = [ivcht11_sh['C6'].value, ivcht11_sh['C7'].value, ivcht11_sh['C8'].value, ivcht11_sh['C9'].value,
              ivcht11_sh['C10'].value]
ivcht_r1_3 = [ivcht11_sh['D6'].value, ivcht11_sh['D7'].value, ivcht11_sh['D8'].value, ivcht11_sh['D9'].value,
              ivcht11_sh['D10'].value]
ivcht_r1_4 = [ivcht11_sh['E6'].value, ivcht11_sh['E7'].value, ivcht11_sh['E8'].value, ivcht11_sh['E9'].value,
              ivcht11_sh['E10'].value]
ivcht_r1_5 = [ivcht11_sh['F6'].value, ivcht11_sh['F7'].value, ivcht11_sh['F8'].value, ivcht11_sh['F9'].value,
              ivcht11_sh['F10'].value]
ivcht_r1_6 = [ivcht11_sh['G6'].value, ivcht11_sh['G7'].value, ivcht11_sh['G8'].value, ivcht11_sh['G9'].value,
              ivcht11_sh['G10'].value]
############################################################################################################
ivcht_r2_1 = [ivcht11_sh['B18'].value, ivcht11_sh['B19'].value, ivcht11_sh['B20'].value, ivcht11_sh['B21'].value,
              ivcht11_sh['B22'].value]
ivcht_r2_2 = [ivcht11_sh['C18'].value, ivcht11_sh['C19'].value, ivcht11_sh['C20'].value, ivcht11_sh['C21'].value,
              ivcht11_sh['C22'].value]
ivcht_r2_3 = [ivcht11_sh['D18'].value, ivcht11_sh['D19'].value, ivcht11_sh['D20'].value, ivcht11_sh['D21'].value,
              ivcht11_sh['D22'].value]
ivcht_r2_4 = [ivcht11_sh['E18'].value, ivcht11_sh['E19'].value, ivcht11_sh['E20'].value, ivcht11_sh['E21'].value,
              ivcht11_sh['E22'].value]
ivcht_r2_5 = [ivcht11_sh['F18'].value, ivcht11_sh['F19'].value, ivcht11_sh['F20'].value, ivcht11_sh['F21'].value,
              ivcht11_sh['F22'].value]
ivcht_r2_6 = [ivcht11_sh['G18'].value, ivcht11_sh['G19'].value, ivcht11_sh['G20'].value, ivcht11_sh['G21'].value,
              ivcht11_sh['G22'].value]
############################################################################################################
while True:
    try:
        week_day = datetime.datetime.today().weekday()
        print("Оно включено 1")
        messages = vk.method("messages.getConversations", {"offset": 0, "count": 20, "filter": "unanswered"})
        if messages["count"] >= 1:
            print("Есть смс")
            id = messages["items"][0]["last_message"]["from_id"]
            body = messages["items"][0]["last_message"]["text"]
            if now_week % 2 == 1:
                if body.lower() == "расписание понедельник" or (body.lower() == "сегодня" and week_day == 0) or (body.lower() == "завтра" and week_day == 6):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r2_1[0]) + pp + p2 + str(
                        ivcht_r2_1[1]) + pp + p3 + str(ivcht_r2_1[2]) + pp + p4 + str(ivcht_r2_1[3]),
                                                "random_id": random.randint(1, 2147483647)})

                elif body.lower() == "расписание вторник" or (body.lower() == "сегодня" and week_day == 1) or (body.lower() == "завтра" and week_day == 0):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r2_2[0]) + pp + p2 + str(
                        ivcht_r2_2[1]) + pp + p3 + str(ivcht_r2_2[2]) + pp + p4 + str(ivcht_r2_2[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание среда" or (body.lower() == "сегодня" and week_day == 2) or (body.lower() == "завтра" and week_day == 1):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r2_3[0]) + pp + p2 + str(
                        ivcht_r2_3[1]) + pp + p3 + str(ivcht_r2_3[2]) + pp + p4 + str(ivcht_r2_3[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание четверг" or (body.lower() == "сегодня" and week_day == 3) or (body.lower() == "завтра" and week_day == 2):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r2_4[0]) + pp + p2 + str(
                        ivcht_r2_4[1]) + pp + p3 + str(ivcht_r2_4[2]) + pp + p4 + str(ivcht_r2_4[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание пятница" or (body.lower() == "сегодня" and week_day == 4) or (body.lower() == "завтра" and week_day == 3):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r2_5[0]) + pp + p2 + str(
                        ivcht_r2_5[1]) + pp + p3 + str(ivcht_r2_5[2]) + pp + p4 + str(ivcht_r2_5[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание суббота" or (body.lower() == "сегодня" and week_day == 5) or (body.lower() == "завтра" and week_day == 4):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r2_6[0]) + pp + p2 + str(
                        ivcht_r2_6[1]) + pp + p3 + str(ivcht_r2_6[2]) + pp + p4 + str(ivcht_r2_6[3]),
                                                "random_id": random.randint(1, 2147483647)})
                else:
                    vk.method("messages.send", {"peer_id": id,
                                                "message": "Команды вводятся по форме: \"расписание день_недели\", например : \"Расписание среда\"",
                                                "random_id": random.randint(1, 2147483647)})
                #################################################################################
            elif now_week % 2 == 0:
                if body.lower() == "расписание понедельник" or (body.lower() == "сегодня" and week_day == 0) or (body.lower() == "завтра" and week_day == 6):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r1_1[0]) + pp + p2 + str(
                        ivcht_r1_1[1]) + pp + p3 + str(ivcht_r1_1[2]) + pp + p4 + str(
                        ivcht_r1_1[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание вторник" or (body.lower() == "сегодня" and week_day == 1) or (body.lower() == "завтра" and week_day == 0):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r1_2[0]) + pp + p2 + str(
                        ivcht_r1_2[1]) + pp + p3 + str(ivcht_r1_2[2]) + pp + p4 + str(
                        ivcht_r1_2[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание среда" or (body.lower() == "сегодня" and week_day == 2) or (body.lower() == "завтра" and week_day == 1):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r1_3[0]) + pp + p2 + str(
                        ivcht_r1_3[1]) + pp + p3 + str(ivcht_r1_3[2]) + pp + p4 + str(
                        ivcht_r1_3[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание четверг" or (body.lower() == "сегодня" and week_day == 3) or (body.lower() == "завтра" and week_day == 2):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r1_4[0]) + pp + p2 + str(
                        ivcht_r1_4[1]) + pp + p3 + str(ivcht_r1_4[2]) + pp + p4 + str(
                        ivcht_r1_4[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание пятница" or (body.lower() == "сегодня" and week_day == 4) or (body.lower() == "завтра" and week_day == 3):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r1_5[0]) + pp + p2 + str(
                        ivcht_r1_5[1]) + pp + p3 + str(ivcht_r1_5[2]) + pp + p4 + str(
                        ivcht_r1_5[3]),
                                                "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "расписание суббота" or (body.lower() == "сегодня" and week_day == 5) or (body.lower() == "завтра" and week_day == 4):
                    vk.method("messages.send", {"peer_id": id, "message": p1 + str(ivcht_r1_6[0]) + pp + p2 + str(
                        ivcht_r1_6[1]) + pp + p3 + str(ivcht_r1_6[2]) + pp + p4 + str(
                        ivcht_r1_6[3]), "random_id": random.randint(1, 2147483647)})
                elif body.lower() == "сегодня":
                    print("Сегодня")
                else:
                    vk.method("messages.send", {"peer_id": id,
                                                "message": "Команды вводятся по форме: \"расписание день_недели\", например : \"Расписание среда\", сегодня - на сегодня, завтра - на завтра", "keyboard" : keyboard,
                                                "random_id": random.randint(1, 2147483647)})
        print("Оно включено 2")
    except Exception as E:
        time.sleep(1)
