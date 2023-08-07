from typing import List

class Authors:
    name: List[str]
    first_name: List[str]
    last_name: List[str]
    no_of_authors: str
    corresponding_author: List[str]

    def __init__(self, author_list: List[int], corresponding_author: List[int] = None):
        if author_list is None:
            return
        self.first_name = []
        self.last_name = []
        self.name = author_list
        self.no_of_authors = len(author_list)
        self.corresponding_author = CorrespondingAuthor(corresponding_author)
        for author in author_list:
            last_name = author.split(' ')[-1]
            while last_name == '' or last_name == 'II' or last_name == 'Alwi':
                author_first = [author.split(' ')[:-1][0]]
                if len(author.split(' ')[:-1]) > 1:
                    for np in author.split(' ')[:-1][1:]:
                        author_first[0] += ' ' + np
                author = author_first[0]
                # author =  author[:-1]
                last_name =  author.split(' ')[-1] + ' ' + last_name if last_name != '' else author.split(' ')[-1] + last_name
            self.last_name.append(last_name)
            first_name = ''
            for name_part in author.split(' ')[:-1]:
                first_name += ' ' + name_part.strip() if len(first_name) > 0 else name_part.strip()
            self.first_name.append(first_name)



class CorrespondingAuthor(Authors):
    affiliation: str
    email: str

    def __init__(self, author_list):
        super().__init__(author_list)