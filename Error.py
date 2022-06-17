class Error:
    """
    Класс для показа ошибок
    """

    def __init__(self, in_text_error: str, in_place_of_error: str) -> None:
        """
        :param in_text_error: описание ошибки
        :param in_place_of_error: показать место ошибки
        """
        self.text_error = in_text_error
        self.place_of_error = in_place_of_error

    def __str__(self):
        """
        :return: Вывод ошибки в строку
        """
        return f"Ошибка: {self.text_error}\n\t{self.place_of_error} <---------\n"