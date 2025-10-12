def is_list_string(input_list: list) -> bool:
    """Checa se todos os elementos da lista `input_list` são tipo texto (str).

    Args:
        input_list (list): A lista que será checada.

    Returns:
        bool: True se todos os elementos forem tipo texto, False se não forem.
    """

    return all(isinstance(item, str) for item in input_list)


if __name__ == "__main__":
    integer_list = ["A", "B", "C", "D"]
    resultado = is_list_string(integer_list)
    print(resultado)
