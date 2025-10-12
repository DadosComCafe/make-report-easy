from abc import ABC, abstractmethod
from typing import List


class AbstractSheet(ABC):
    def __init__(self, path: str): ...

    @abstractmethod
    def get_uniquetype_columns(self) -> List: ...

    @abstractmethod
    def create_unique_type_sheet(self) -> None: ...
