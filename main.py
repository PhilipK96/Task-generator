import random as rnd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx2pdf import convert

def create_vars(group: str, date_of_issue: str, students):
    i = -1

    for name in students: 
        i = i + 1   
        # Нагрузка 110 кВ (только для ПС 1)
        P1_110_max = rnd.randrange(30,40,1)
        tg1_110_max = round(rnd.uniform(0.2,0.5),2)
        P1_110_min = round((rnd.randint(80,90))/100 * P1_110_max)
        tg1_110_min = round((rnd.randint(90,100))/100 * tg1_110_max,2)
        # Нагрузки 10 кВ (все подстанции)
        P1_10_max = rnd.randrange(30,40,1)
        P2_10_max = rnd.randrange(8,40,1)
        P3_10_max = rnd.randrange(8,40,1)
        P4_10_max = rnd.randrange(8,40,1)

        tg1_10_max = round(rnd.uniform(0.2,0.5),2)
        tg2_10_max = round(rnd.uniform(0.2,0.5),2)
        tg3_10_max = round(rnd.uniform(0.2,0.5),2)
        tg4_10_max = round(rnd.uniform(0.2,0.5),2)

        P1_10_min = round((rnd.randint(30,50))/100 * P1_10_max)
        P2_10_min = round((rnd.randint(30,50))/100 * P2_10_max)
        P3_10_min = round((rnd.randint(30,50))/100 * P3_10_max)
        P4_10_min = round((rnd.randint(30,50))/100 * P4_10_max)

        tg1_10_min = round((rnd.randint(90,100))/100 * tg1_10_max,2)
        tg2_10_min = round((rnd.randint(90,100))/100 * tg2_10_max,2)
        tg3_10_min = round((rnd.randint(90,100))/100 * tg3_10_max,2)
        tg4_10_min = round((rnd.randint(90,100))/100 * tg4_10_max,2)
        # Время максимальных нагрузок
        Tma_1 = rnd.randrange(2000,7000,50)
        Tma_2 = rnd.randrange(2000,7000,50)
        Tma_3 = rnd.randrange(2000,7000,50)
        Tma_4 = rnd.randrange(2000,7000,50)

        tg_A_max = round(rnd.uniform(0.25,0.5),2)
        tg_A_min = round((rnd.randint(90,100))/100 * tg_A_max,2)

        Km = round(rnd.uniform(0.4,0.9),2)
        h = rnd.randrange(4,9,1)
        alpha = (rnd.randrange(3,8,1))/10
        # Задание списка ОЭС и выбор какой-то одной для варианта
        OES_list = ('Востока','Сибири','Урала','Средней Волги','Юга','Центра','Северо-запада')
        OES = rnd.choice(OES_list)

        if OES == 'Востока':
            T_winter = -round(rnd.uniform(20,22),1)
        elif OES == 'Сибири':
            T_winter = -round(rnd.uniform(21,23),1)
        elif OES == 'Урала':
            T_winter = -round(rnd.uniform(18,20),1)
        elif OES == 'Средней Волги':
            T_winter = -round(rnd.uniform(9,11),1)
        elif OES == 'Юга':
            T_winter = -round(rnd.uniform(4,6),1)
        elif OES == 'Центра':
            T_winter = -round(rnd.uniform(10,12),1)
        elif OES == 'Северо-запада':
            T_winter = -round(rnd.uniform(3,5),1)
        

        T_p = rnd.randint(3,5)
        En = (rnd.randrange(70,100,1))/10
        price = (rnd.randrange(24,36,1))/10
        drawing_scale = rnd.randrange(8,12)

        # Запись
        doc = DocxTemplate("./src/template.docx")
        context = {'student_name': str(students[i]),    
            'variant_number': i+1,
            'P1_110_max': P1_110_max,
            'tg1_110_max': tg1_110_max,
            'P1_110_min': P1_110_min,
            'tg1_110_min': tg1_110_min,
            'P1_10_max': P1_10_max,
            'P2_10_max': P2_10_max,
            'P3_10_max': P3_10_max,
            'P4_10_max': P4_10_max,
            'tg1_10_max': tg1_10_max,
            'tg2_10_max': tg2_10_max,
            'tg3_10_max': tg3_10_max,
            'tg4_10_max': tg4_10_max,
            'P1_10_min': P1_10_min,
            'P2_10_min': P2_10_min,
            'P3_10_min': P3_10_min,
            'P4_10_min': P4_10_min,
            'tg1_10_min': tg1_10_min,
            'tg2_10_min': tg2_10_min,
            'tg3_10_min': tg3_10_min,
            'tg4_10_min': tg4_10_min,
            'Tma_1': Tma_1,
            'Tma_2': Tma_2,
            'Tma_3': Tma_3,
            'Tma_4': Tma_4,
            'tg_A_max': tg_A_max,
            'tg_A_min': tg_A_min,
            'Km': Km,
            'h': h,
            'alpha': alpha,
            'OES': OES,
            'T_winter': T_winter,
            'T_p': T_p,
            'En': En,
            'price': price,
            'drawing_scale': drawing_scale,
            
            'Topology':InlineImage(doc,('./src/'+str(rnd.randint(1,4))+'.png'),height=Mm(65)),
            'date_of_issue':date_of_issue,
            'group': group
        }
       
        doc.render(context)
        doc.save("./output/Вариант_"+ str(i+1) +"_" + str(students[i] + ".docx"))
        # Конвертация в pdf
        convert("./output/Вариант_"+ str(i+1) +"_" + str(students[i] + ".docx"),"./output/Вариант_"+ str(i+1) +"_" + str(students[i] + ".pdf"))



if __name__ == "__main__":
    students = (
    "Иванов И.И.",
    "Алексеев А.А",
    "Петров А.А.",						
    )
create_vars("3.131", "26.01.2022", students)
