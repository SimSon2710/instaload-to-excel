import lzma
import json
from os.path import join, basename, exists, dirname
from os import mkdir
from datetime import datetime
from openpyxl.drawing.image import Image
from PIL import Image as PImage
from InstaFile import InstaFile


class InstaPost:
    all_posts = {}
    comments_count = 0

    def __init__(self, insta_file: InstaFile):
        self.insta_file = insta_file
        self.meta_data = None
        self.username = None
        self.text = None
        self.jpg_path_list = insta_file.jpg_path_list
        self.post_is_part_of_collection = insta_file.is_collection
        self.image_list = []
        self.post_time = None
        self.post_likes = None
        self.post_id = None
        self.profile_full_name = None
        self.profile_category = None
        self.profile_biography = None
        self.profile_followed_by = None
        self.profile_follows = None
        self.profile_is_business_account = None
        self.profile_is_joined_recently = None
        self.profile_external_url = None

        self.comment_data = None
        self.comments = {}

        self.row_height = 185

    def json_to_meta_data(self):
        if self.insta_file.json_meta_path:
            with lzma.open(self.insta_file.json_meta_path) as f:
                json_bytes = f.read()
                stri = json_bytes.decode('utf-8')
                self.meta_data = json.loads(stri)
            self.meta_data.pop("instaloader", None)

    def json_to_comment_data(self):
        if self.insta_file.json_comments_path:
            with open(self.insta_file.json_comments_path) as f:
                self.comment_data = json.load(f)
            InstaPost.comments_count += 1

    @classmethod
    def img_rescale(cls, orig_size_width, orig_size_height, aimed_img_width, max_img_height):
        factor_height = max_img_height / orig_size_height
        factor_weight = aimed_img_width / orig_size_width
        if factor_height <= factor_weight:
            return factor_height
        elif factor_weight < factor_height:
            return factor_weight
        else:
            return False

    def check_long_text(self):
        standard_height = self.row_height
        long_txt_factor = 2.3
        long_txt_resize_factor = 0.4
        if self.text or self.profile_biography:
            len_bio = len(self.profile_biography)
            len_text = len(self.text)
            n_text_bio = self.profile_biography.count("\n")
            n_text = self.text.count("\n")
            if n_text_bio > 0:
                len_bio = len_bio + n_text_bio * 35
            if n_text > 0:
                len_text = len_text + n_text * 35
            if len_bio > 0 or len_text > 0:
                if (len_text - len_bio) > 0 and len_text > standard_height * long_txt_factor:
                    self.row_height = len_text * long_txt_resize_factor
                elif (len_text - len_bio) < 0 and len_bio > standard_height * long_txt_factor:
                    self.row_height = len_bio * long_txt_resize_factor

    def make_img(self):
        for img_path in self.jpg_path_list:
            img_pil = PImage.open(img_path)
            aimed_img_width = 190
            max_img_height = 240
            if not exists(join(dirname(img_path), "thumbnails")):
                mkdir(join(dirname(img_path), "thumbnails"))
            thumbnail_path = join(dirname(img_path), "thumbnails", f"{basename(img_path)[:-4]}_thumbnail.jpg")
            img_pil.save(thumbnail_path, quality=10)
            img = Image(thumbnail_path)
            img_width, img_height = PImage.open(thumbnail_path).size
            if InstaPost.img_rescale(img_width, img_height, aimed_img_width, max_img_height):
                img.width = img.width * InstaPost.img_rescale(img_width, img_height, aimed_img_width, max_img_height)
                img.height = img.height * InstaPost.img_rescale(img_width, img_height, aimed_img_width,
                                                                max_img_height)
            self.image_list.append(img)

    def pipe_meta_data(self):
        meta = self.meta_data
        if meta:
            while True:
                try:
                    meta["node"]["edge_media_preview_like"]["count"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.username = meta["node"]["owner"]["username"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_full_name = meta["node"]["owner"]["full_name"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_is_business_account = meta["node"]["owner"]["is_business_account"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_is_joined_recently = meta["node"]["owner"]["is_joined_recently"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_category = meta["node"]["owner"]["category_enum"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_external_url = meta["node"]["owner"]["external_url"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.text = meta["node"]["edge_media_to_caption"]["edges"][0]["node"]["text"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.post_time = \
                        datetime.utcfromtimestamp(int(meta["node"]["taken_at_timestamp"])).strftime(
                            "%Y-%m-%d %H:%M:%S")
                except (KeyError, IndexError):
                    pass
                try:
                    self.post_likes = meta["node"]["edge_media_preview_like"]["count"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_biography = meta["node"]["owner"]["biography"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_followed_by = meta["node"]["owner"]["edge_followed_by"]["count"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.profile_follows = meta["node"]["owner"]["edge_follow"]["count"]
                except (KeyError, IndexError):
                    pass
                try:
                    self.post_id = meta["node"]["id"]
                except (KeyError, IndexError):
                    pass
                break

    def pipe_comment_data(self):
        comments = self.comment_data
        if comments:
            self.comments["post_id"] = self.post_id
            self.comments["post_of_user"] = self.username
            self.comments["comments"] = comments

    def add_to_all_posts(self):
        for image in self.image_list:
            self.all_posts[image] = self
