from fuzzywuzzy import fuzz
import numpy as np

johor = ['Segamat','Tangkak','Muar','Batu Pahat','Kluang','Mersing','Kota Tinggi','Kulai','Pontian','Johor Bahru']
kedah = ['Kubang Pasu','Padang Terap','Pokok Sena','Kota Setar','Pendang','Yan','Sik','Kuala Muda','Baling','Kulim','Bandar Baharu']
kelantan = ['Tumpat','Pasir Mas','Kota Bharu','Bachok','Pasir Puteh','Machang','Tanah Merah','Jeli','Kuala Krai','Gua Musang']
melaka = ['Alor Gajah','Jasin','Melaka Tengah']
negeri_sembilan = ['Jelebu','Jempol','Kuala Pilah','Seremban','Rembau','Port Dickson','Tampin']
pahang = ['Cameron Highland','Lipis','Jerantut','Kuantan','Raub','Temerloh','Maran','Pekan','Bera','Bentong','Rompin']
perak = ['Hulu Perak','Kerian','Kuala Kangsar','Kinta','Kampar','Batang Padang','Perak Tengah','Manjung','Hilir Perak','Muallim','Larut,Matang&Selama']
perlis = ['Padang Besar','Kangar','Arau']
pulau_pinang = ['Seberang Perai Utara','Seberang Perai Tengah','Seberang Perai Selatan','Timur Laut','Barat Daya']
sabah = ['Pantai Barat (Utara)','Pantai Barat (Selatan)','Sandakan','Beaufort','Sandakan','Lahad Datu','Tawau','Keningau']
sarawak = ['Limbang','Miri','Bintulu','Kapit','Sibu','Mukah','Sarikei','Betong','Sri Aman','Samarahan','Kuching','Seriah']
selangor = ['Sabak Bernam','Hulu Selangor','Kuala Selangor','Gombak','Klang','Petaling','Hulu Langat','Kuala Langat','Sepang']
terengganu = ['Besut','Setiu','Kuala Nerus','Kuala Terengganu','Marang','Hulu Terengganu','Dungun','Kemaman']
state = ['Johor','Kedah','Kelantan','Melaka','Negeri Sembilan','Pahang','Perak','Perlis','Pulau Pinang','Sabah','Sarawak','Selangor','Terengganu']

def check(word,arr):
  length = len(arr)
  similarity = {}
  x = 0
  while x < length:
    ration = fuzz.ratio(word,arr[x])
    similarity[arr[x]]=ration
    x += 1
  word = (max(similarity, key=similarity.get))
  return word

def spell_state(word):
  word = str(word)
  word = check(word,state)
  return word

def spell(word,page):
  word = str(word)
  page = int(page)
  state = ''
  if page in range(12,40):
    word=check(word,johor)
    state = 'Johor'
  if page in range(54,63):
    word=check(word,kedah)
    state = 'Kedah'
  if page in range(66,78):
    word=check(word,kelantan)
    state = 'Kelantan'
  if page in range(82,88):
    word=check(word,melaka)
    state = 'Melaka'
  if page in range(92,108):
    word=check(word,negeri_sembilan)
    state = 'Negeri Sembilan'
  if page in range(112,152):
    word=check(word,pahang)
    state = 'Pahang'
  if page in range(156,188):
    word=check(word,perak)
    state = 'Perak'
  if page == 192:
    word=check(word,perlis)
    state = "Perlis"
  if page in range(196,197):
    word=check(word,pulau_pinang)
    state = 'Pulau Pinang'
  if page in range(200,305):
    word=check(word,sabah)
    state = 'Sabah'
  if page in range(308,357):
    word=check(word,sarawak)
    state = 'Sarawak'
  if page in range(360,369):
    word=check(word,selangor)
    state = 'Selangor'
  if page in range(372,382):
    word=check(word,terengganu)
    state = 'Terengganu'

  return word,state
