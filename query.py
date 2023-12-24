import mysql.connector
from tkinter import END


def create_worker(person, profession):
    cnt = 1
    for _ in profession:
        cnt += 0.5

    profession.append(person.position)
    profession.append(person.work_exp)

    cnx = mysql.connector.connect(user='root', password='root',port='3306',
                                  host='localhost', database='practica',
                                  )

    cursor = cnx.cursor()

    query = "INSERT INTO person (full_name, Date, department, family_status," \
            " position, work_exp, count_position) VALUES (%s, %s, %s, %s, %s, %s, %s)"

    message = (person.full_name, person.Date, person.departament, person.family_status, person.position, person.work_exp, cnt)
    cursor.execute(query, message)

    cnx.commit()
    cursor.close()
    cnx.close()

    add_proffesion(profession, person.full_name)


def add_proffesion(professions, person_name):
    cnt = 0
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                  host='localhost', database='practica'
                                  )
    cursor = cnx.cursor()

    query = "SELECT id FROM person WHERE full_name = %s"

    cursor.execute(query, (person_name,))

    result = cursor.fetchone()

    for i in range(0 ,len(professions), 2):
        query = "INSERT INTO professions (profession, workerID, work_exp) VALUES (%s, %s, %s)"
        message = (professions[i], result[0], professions[i+1])
        cursor.execute(query, message)
        cnx.commit()

        cnt += 2
        if cnt == 10:
            break

    cursor.close()
    cnx.close()


def check_person(full_name):
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                  host='localhost', database='practica'
                                  )
    cursor = cnx.cursor()

    query = "SELECT id FROM person WHERE full_name = %s"

    cursor.execute(query, (full_name,))

    result = cursor.fetchone()

    if result is not None:
        return True
    else:
        return False


def deleteEmployee(full_name):
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                  host='localhost', database='practica')

    cursor = cnx.cursor()

    query = "SELECT id FROM person WHERE full_name = %s"
    cursor.execute(query, (full_name,))
    result = cursor.fetchone()

    delete_query = "DELETE FROM professions WHERE workerID = %s"
    cursor.execute(delete_query, (result[0],))

    delete_person_query = "DELETE FROM person WHERE full_name = %s"
    cursor.execute(delete_person_query, (full_name,))
    cnx.commit()

    cursor.close()
    cnx.close()


def search_worker_prof():
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                  host='localhost', database='practica')

    cursor = cnx.cursor()

    sql = "SELECT full_name, count_position FROM person"

    cursor.execute(sql)

    result = cursor.fetchall()

    cursor.close()
    cnx.close()

    return result


def search_all_worker_prof(family_status, profesion):
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                  host='localhost', database='practica')

    cursor = cnx.cursor()

    sql = '''
    SELECT p.full_name AS 'Ф. И. О. сотрудника', pr.work_exp AS 'Стаж работы по этой профессии'
    FROM person p
    JOIN professions pr ON p.id = pr.workerID
    WHERE p.family_status = %s 
    AND pr.Profession = %s
    '''

    cursor.execute(sql,(family_status, profesion))

    result = cursor.fetchall()

    cursor.close()
    cnx.close()

    return result


def change_department_person(full_name, department):
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                  host='localhost', database='practica')

    cursor = cnx.cursor()

    query = ("UPDATE person SET department=%s"
             "WHERE full_name=%s")

    cursor.execute(query, (department, full_name))

    cnx.commit()
    cursor.close()
    cnx.close()


def select_full_name_where_department(department):
    cnx = mysql.connector.connect(user='root', password='root', port='3306',
                                 host='localhost', database='practica')

    cursor = cnx.cursor()

    query = ("SELECT full_name FROM person WHERE department = %s")

    cursor.execute(query, (department,))

    result = cursor.fetchall()

    cursor.close()
    cnx.close()

    return result


def delete_full_name_where_department(full_name_list):
    for full_name in full_name_list:
        deleteEmployee(full_name[0])


def updateTable(tree):
    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="practica",
        port='3306'
    )

    mycursor = mydb.cursor()

    mycursor.execute("SELECT * FROM person")
    updateResult = mycursor.fetchall()

    for i in tree.get_children():
        tree.delete(i)
    for row in updateResult:
        tree.insert("", END, values=(row[1], row[2], row[3], row[4], row[5], row[6], row[7]))

    mydb.close()
    mycursor.close()


def create_table():
    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="practica",
        port='3306'
    )

    mycursor = mydb.cursor()

    mycursor.execute("SELECT * FROM person")

    result = mycursor.fetchall()

    return result


def find_all_profession_person(full_name):
    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="practica",
        port='3306'
    )

    cursor = mydb.cursor()

    query = ("SELECT p.ID, pr.Profession \
            FROM practica.person p \
            JOIN practica.professions pr ON p.ID = pr.workerID \
            WHERE p.full_name = %s;")

    cursor.execute(query, (full_name,))

    result = cursor.fetchall()

    return result


def search_all_full_name():
    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="practica",
        port='3306'
    )

    cursor = mydb.cursor()

    query = ("SELECT full_name FROM person")

    cursor.execute(query)

    result = cursor.fetchall()

    return result


def search_all_department():

    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="practica",
        port='3306'
    )

    cursor = mydb.cursor()

    query = ("SELECT department FROM person")

    cursor.execute(query)

    result = cursor.fetchall()

    return result