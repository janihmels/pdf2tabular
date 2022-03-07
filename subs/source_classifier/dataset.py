
from torch.utils.data import Dataset
from utils import get_data, organize_data, augment_name
from transformers import BertTokenizer
import random
import torch
from utils import prepro

MAX_LENGTH = 100


class ds(Dataset):

    def __init__(self, data, augment):
        super().__init__()

        self.tokenizer = BertTokenizer.from_pretrained('prajjwal1/bert-tiny')
        self.data = data

        # -- preprocess --

        self.classes_start_indexes = []
        self.data_list = []

        cls_start_idx = 0
        for cls_sources, cls_label in self.data:
            self.classes_start_indexes.append((cls_start_idx, cls_label))
            cls_start_idx += len(cls_sources)

            self.data_list += [(source, cls_label) for source in cls_sources]

        self.augment = augment

    def __len__(self):
        return len(self.data_list)

    def __getitem__(self, idx):
        name, label = self.data_list[idx]

        if self.augment:
            # augment the data
            if random.randint(0, 1) == 1:
                name = augment_name(name)

        name = prepro(self.tokenizer, name)

        return torch.Tensor(name['input_ids']).long(), torch.Tensor(name['attention_mask']).long(), \
               torch.Tensor(name['token_type_ids']).long(), torch.Tensor([label]).long()
