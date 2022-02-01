from os.path import exists, join, basename
from os import listdir
from GUI import GUI


class InstaFile:
    meta_ext = ".json.xz"

    def __init__(self):
        self.jpg_path_list = []
        self.json_meta_path = ""
        self.json_comments_path = ""
        self.is_collection = False

    def parse_json_path(self, file_path):
        self.json_meta_path = file_path

    def get_jpg_files(self):
        base_path = self.json_meta_path[:-8]
        if not exists(f"{base_path}_1.jpg"):
            if exists(f"{base_path}.jpg"):
                self.jpg_path_list.append(f"{base_path}.jpg")
                GUI.jpg_counter += 1
        else:
            if exists(f"{base_path}_1.jpg"):
                self.jpg_path_list.append(f"{base_path}_1.jpg")
                self.is_collection = True
                GUI.jpg_counter += 1
                num = 2
                while True:
                    if exists(f"{base_path}_{str(num)}.jpg"):
                        self.jpg_path_list.append(f"{base_path}_{str(num)}.jpg")
                        GUI.jpg_counter += 1
                        num += 1
                    else:
                        break

    def get_comments(self):
        comments_path = f"{self.json_meta_path[:-8]}_comments.json"
        if exists(comments_path):
            self.json_comments_path = comments_path


if __name__ == '__main__':
    for folder in listdir(".\loads"):
        for file in listdir(join(".\loads", folder)):
            Post = InstaFile()
            Post.parse_json_path(join(".\loads", folder, file))
            Post.get_jpg_files()
            if Post.jpg_path_list and Post.json_meta_path:
                print(Post.is_collection, "--->", basename(Post.json_meta_path), ": ", Post.jpg_path_list)
