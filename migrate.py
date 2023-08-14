import sqlite3
import bcrypt
db = sqlite3.connect('database.db')
password = 'STr0nG_P@SSW0rD'
username = 'admin'
cursor = db.cursor()


cursor.execute('''
    CREATE TABLE users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        username TEXT,
        password TEXT
    )
''')

cursor.execute('''
    CREATE TABLE tokens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        token TEXT,
        FOREIGN KEY(user_id) REFERENCES users(id)
    )
''')
           
hashed_password = bcrypt.hashpw(password.encode('utf8'), bcrypt.gensalt())
cursor.execute('INSERT INTO users (username, password) VALUES (?, ?)', 
           (username, hashed_password))


print('Database and tables created')

db.commit()
cursor.close()
db.close()