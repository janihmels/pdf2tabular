
from torch.utils.data import Dataset
from utils import get_data
from transformers import BertTokenizer
import random
import torch
from utils import prepro

MAX_LENGTH = 100


class ds(Dataset):

    def __init__(self):
        super().__init__()
        self.names, self.not_names = get_data(featurize_data=False)

        tokenizer = self.tokenizer = BertTokenizer.from_pretrained('prajjwal1/bert-tiny')

        # -- preprocessing

        self.names = [prepro(tokenizer, name) for name in self.names]
        self.not_names = [prepro(tokenizer, not_name) for not_name in self.not_names]

    def __len__(self):
        return len(self.names) + len(self.not_names)

    def __getitem__(self, idx):

        if random.uniform(0, 1) > 0.5:
            name, label = self.names[idx % len(self.names)], 1
        else:
            name, label = self.not_names[idx % len(self.not_names)], 0

        return torch.Tensor(name['input_ids']).long(), torch.Tensor(name['attention_mask']).long(), \
               torch.Tensor(name['token_type_ids']).long(), torch.Tensor([label]).float()
