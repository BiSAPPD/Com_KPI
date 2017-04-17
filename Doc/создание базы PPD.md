# **Создание базы PPD для e-academie v3**

## **Подготовка скриптов**

### **Этап 1-й. Стандартизация адресов**
В базах e-academie марок **65 000** записей (клиентов)

| марка | количество записей|
|---- |:----:|
|LP|15 550|
|MX|42 300|
|KR|1 600|
|RD|2 300|

Все написания адресов были стандартиризованы.

|старое написание| |стандартизированное написание |
|---|---|---|
|Москва Ленина 15  | -> | г Москва ул Ленина д 15 |
|город Пермь улица Ленина 28 |->| г Пермь ул Ленина д 28 |


Также, все адреса были обогащены географической информацией

|тип | данные |
|---|---|
|	Полный Адрес	|	г Пермь, ул Ленина, д 28	|
|	Индекс	|	614000	|
|	Страна	|	Россия	|
|	Регион	|	Пермский край	|
|	Город	|	г Пермь	|
|	Район города	|	Ленинский район	|
|	Улица	|	ул Ленина	|
|	Дом	|	д 28	|
|	Код КЛАДР	|	5900000100006330000|
|	Код ФИАС	|	13b3cccb-3d26-44e8-94f0-e652b12782d7	|
|	Уровень по ФИАС	|	8: дом	|
|	Признак центра района или региона	|	2:  центр региона |
|	Часовой пояс	|	UTC+5	|
|	Широта	|	58.0142443	|
|	Долгота	|	56.2488834	|

### **Этап 2-й. Унификация названий клинетов**

В одной базе марки, может быть дубликат клиента с разным написанием названия.<br>
В базах разных брендах, один и тот же клиент может иметь отличающиеся в написани названия.

К примеру, по адресу: г Пермь б-р Гагарина, мы нашли 5 салонов с разными названиями. 

||исходное название| адрес||унифицированное||
|---|---|---|----|---|---|
|ES|Графиня Галакрисо| б-р Гагарина, д 65А/1 |**->** | Графиня Галакрисо| б-р Гагарина, д 65А |
|LP|Графиня Галакрисо, ИП Пепеляева Л. А.| б-р Гагарина, д 65А |**->**|Графиня Галакрисо| б-р Гагарина, д 65А |
|MX|Galakriso| б-р Гагарина, д 65А |**->**|Графиня Галакрисо| б-р Гагарина, д 65А |
|KR|Графиня| б-р Гагарина, д 65 |**->**|Графиня Галакрисо| б-р Гагарина, д 65А |
|RD|Галакрисо| б-р Гагарина, д 65 |**->**|Графиня Галакрисо| б-р Гагарина, д 65А |

Всего было унифицировано 7 500 различных написаний названий по каждой улице. В примере выше - одно унифицирование.

### **Этап 3-й. Привязка клиентов к коммерческой структуре**

База клинетов состоит из потенциальных и активных клинетов.
На основе информации - *"какой представитель работает с наибольшим количеством точек по данной улице в этом городе"*, каждый адрес потенциальных клинетов был привязан к коммерческому сектору.

## **Резюме**

Для автоматического поиска дубликатов или объединения клиентов из разных баз, необходимо абсолютное идентичное написание: названия клинета, адреса.
Также для увеличения критериев сопоставления, мы используем: телефон салона, емейл, данные о руководителе салона.


## **Проверка скриптов**

### **Проверка городов**

Проводил проверку по отработанному скрипту по городам [ссылка](https://drive.google.com/file/d/0BxVHA4PO8GTjYnRUbUV0TUZlNE0/view?usp=drivesdk). 
Для проверки использовал SQL запрос

```
select distinct city_name_geographic, count(city_name_geographic) as cnt
from salons as sln

group by city_name_geographic
order by  cnt desc, city_name_geographic
```
Результат - при беглом анализе данные обработаны

### **Проверка приведения географических регионов**

Проводил проверку по скрипту [ссылка](https://drive.google.com/file/d/0BxVHA4PO8GTjVlpicXF0MGtsQTQ/view?usp=drivesdk). 
Для проверки использовал SQL запрос

```
select  distinct region_name_geographic , city_name_geographic, count(city_name_geographic) as cnt
from salons as sln

group by region_name_geographic , city_name_geographic
order by  region_name_geographic , cnt desc, city_name_geographic
```

### **Проверка привязки географических регионов к Мегарегионам**

Скрипт [ссылка](https://drive.google.com/file/d/0BxVHA4PO8GTjZG42OWtlOUR5ZzA/view?usp=drivesdk) использует логику привязки ком.мегарегиона на основание связки географического региона и города.

По субъективным признакам, достаточна привязка мегарегиона к географическому региону.

При анализе SQL запроса

```
select  distinct region_name_geographic, com_mreg , city_name_geographic,
count(city_name_geographic) as cnt1, count(city_name_geographic) as cnt
from salons as sln

group by region_name_geographic , com_mreg, city_name_geographic
order by  region_name_geographic , com_mreg
```

было выявлено наличие двух проблем:
* отсутсвие у города мегарегиона
* не корректная связка *город + регион и ком.мегарегиона*

Для насыщения базы, был написан SQL запрос, который не дал результатов

```
with educater_hist as 
(select distinct 
(Case when smr.technolog_id = smr.partimer_id then  concat(smr.technolog_id) else Concat(smr.technolog_id , smr.partimer_id)  end )  as educater, 
sln.com_mreg as mreg, 
count(*) over (partition by Concat(smr.technolog_id , smr.partimer_id, sln.com_mreg)) as cbt
from seminars as smr 
left join seminar_users as smu ON smr.id = smu.seminar_id
left join users as usr ON smu.user_id = usr.id
left join salons as sln ON sln.id = usr.Salon_id or sln.salon_manager_id = usr.id

where char_length(sln.city_name_geographic) > 4

order by educater, cbt desc) ,

educater as (
select (Case when smr.technolog_id = smr.partimer_id then  concat(smr.technolog_id) else Concat(smr.technolog_id , smr.partimer_id)  end ) as edu,


(select eh.mreg
from educater_hist as eh
where concat(smr.technolog_id) = eh.educater or concat(smr.partimer_id) = eh.educater
order by eh.cbt desc
limit 1)


from seminars as smr

)
, educater_mreg as 
(select *
from educater
where mreg is Not Null)




select  distinct sln.region_name_geographic, 
(case when sln.com_mreg is Null then edm.mreg else sln.com_mreg end), sln.city_name_geographic,
count(sln.city_name_geographic) as cnt1, count(sln.city_name_geographic) as cnt
from salons as sln
left join users as usr ON sln.id = usr.Salon_id or sln.salon_manager_id = usr.id
left join seminar_users as smu ON smu.user_id = usr.id
left join seminars as smr ON smr.id = smu.seminar_id
left join educater_mreg as edm  ON concat(edm.edu) = concat(smr.technolog_id) or concat(edm.edu) = concat(smr.partimer_id)

group by sln.region_name_geographic , sln.com_mreg, edm.mreg, sln.city_name_geographic
order by  sln.region_name_geographic 
```

В написание была допущена логическая ошибка - делалась таблица Тренер - Регион, а надо было Салон Регион.

### **Проверка обработки написания улиц**

Скрипт [ссылка](https://drive.google.com/file/d/0BxVHA4PO8GTjNTVuVk11SEdUejg/view?usp=drivesdk) используется для стандартизации написания улиц.

По субъективным признакам, улицы приведены в порядок.

При анализе использовался SQL запрос


```
select distinct street, city_name_geographic, count(city_name_geographic) as cnt
from salons as sln

where street like '%Зеленый%'
group by street, city_name_geographic
order by  cnt desc
```

### **Проверка заполнения столбца номера дома**

Скрипт [ссылка](https://drive.google.com/file/d/0BxVHA4PO8GTjVUNjUDhSLXJoV3M/view?usp=drivesdk) используется для заполнения номеров домов, по адресам с отсутсвующей информацией.

По субъективным признакам, дома были прописаны.

При анализе использовался SQL запрос

```
select distinct address, house, street, city_name_geographic, count(city_name_geographic) as cnt
from salons as sln

where city_name_geographic like '%Курган%' and address  like '%31%' 
group by address, house,  street, city_name_geographic
order by  cnt desc
```

### **Проверка изменения написания названий салонов**

Скрипт [ссылка](https://drive.google.com/file/d/0BxVHA4PO8GTjcnRQQjgtTVNXY1E/view?usp=drivesdk) использовался для стандартизации названий салонов.

По субъективным признакам, дома были прописаны.

При анализе использовался SQL запрос

```
select name, city_name_geographic
from salons as sln
where name like '%Solo%' 
```



|рейтинг|`город`|`улица`|дом|название|-|тел.салона|емейл.салона|-|ФИО менеджера|тел.менеджера|емейл менеджера|
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
|1|X|X|X|X||OR|OR|||||
|1|X|X|X|X|||||OR|OR|OR|
|1.5|X|X|X|||OR|OR|||||
|1.5|X|X||X||OR|OR|||||
|1.5|X|X|X||||||OR|OR|OR|
|1.5|X|X||X|||||OR|OR|OR|


