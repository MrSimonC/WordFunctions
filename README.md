# Word Helper
This is a Python class which provides common Microsoft Word functions using the Win32com API.

Example use:
```python
import word_com.py

# Example
from custom_modules.word_com import Word
word = Word()
word.open_word_file('c:\test.xlsx')
# search for "Name:" in word in a table, return the column to the right
print(word.find_table_content("Name:", 1))
word.close()

```