class Student():
    def __init__(self, name, student_id, age, gender):
        self.__name = name
        self.__id = student_id
        self.__age = age
        self.__gender = gender
        self.__grade = 0
        
    def set_grade(self, grade):
        self.__grade = grade
        
    def get_grade(self):
        return self.__grade
        
    def display_student_info(self):
        return "student name: {:s}, student id: {:s}, age: {:d}, gender: {:s}, grade:{:d}"\
                .format(str(self.__name), str(self.__id), self.__age, str(self.__gender), self.__grade)
