import pandas as pd
from PyQt6.QtCore import Qt, QAbstractTableModel, QVariant

class PandasModel(QAbstractTableModel):
    """A model to interface a Qt view with a Pandas DataFrame."""
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return len(self._df.index)

    def columnCount(self, parent=None):
        return len(self._df.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return QVariant()

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._df.iat[index.row(), index.column()]
            return str(value) if pd.notna(value) else ""
        return QVariant()

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return QVariant()

        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section])
        elif orientation == Qt.Orientation.Vertical:
            return str(self._df.index[section])
        return QVariant()
    