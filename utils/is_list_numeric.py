def is_list_numeric(input_list: list) -> bool:
    """Checa se todos os elementos da lista `input_list` são numéricos (integers ou floats).

    Args:
        input_list (list): A lista que será checada.

    Returns:
        bool: True se todos os elementos forem numéricos, False se não forem.
    """

    return all(isinstance(item, (int, float)) for item in input_list)


if __name__ == "__main__":
    integer_list = [1, 2, 3, 4]
    resultado = is_list_numeric(integer_list)
    print(resultado)
