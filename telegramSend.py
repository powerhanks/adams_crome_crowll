
import telegram
import json
from collections import OrderedDict

def telebot(text):
	with open('./configData.json', encoding="utf-8") as data_file:    
	   data = json.load(data_file, object_pairs_hook=OrderedDict)
	#import telegram
	# 토큰을 지정해서 bot을 선언
	bot = telegram.Bot(token=data["telegram_token"])
	# 우선 테스트 봇이니까 가장 마지막으로 bot에게 말을 건 사람의 id를 지정해줄게요.
	# 만약 IndexError 에러가 난다면 봇에게 메시지를 아무거나 보내고 다시 테스트해보세요.
	#chat_id = bot.getUpdates()[-1].message.chat.id
	chat_id=data["telegram_chat_id"]

	bot.sendMessage(chat_id=chat_id,text=text)

if __name__ == "__main__": 
	messageTxt = '테스트 메세지입니다'
	telebot(messageTxt)