
import torch
from torch import nn
import pytorch_lightning as pl
from collections import defaultdict
from transformers import AutoModel, BertTokenizer
from subs.names_classifier.utils import prepro


class NamesClassifier(pl.LightningModule):

    def __init__(self):

        super().__init__()

        self.bert = AutoModel.from_pretrained('prajjwal1/bert-tiny')
        self.final_layer = nn.Linear(in_features=128, out_features=1)
        self.criterion = torch.nn.BCEWithLogitsLoss()

        self.tokenizer = BertTokenizer.from_pretrained('prajjwal1/bert-tiny')

        self.best_val_acc = -1

    def forward(self, input_ids, attention_masks, token_types_masks):

        """
        :param input_ids: Tensor of shape (batch_size, 512, size(vocabulary)).
        :return: Tensor of shape (batch_size,) containing classification logit for each name in the batch.
        """

        _, cls_token = self.bert(input_ids,
                                 attention_mask=attention_masks,
                                 token_type_ids=token_types_masks,
                                 return_dict=False)

        # cls_token.shape = (batch_size, 128)

        return self.final_layer(cls_token).squeeze(dim=1)

    def predict(self, input_ids, attention_masks, token_types_masks):
        """
        :return: A tensor of size (batch_size, ) containing the prediction of the network (1-title, 0-non title)
                for every item in the batch.
        """

        logits = self(input_ids, attention_masks, token_types_masks)
        return logits > 0

    def step(self, batch, batch_idx):
        input_ids, attention_masks, token_types_masks, is_title = batch
        logits = self(input_ids, attention_masks, token_types_masks)  # shape = (batch_size, )

        is_title = is_title.squeeze(dim=1)

        return {'loss': self.criterion(logits, is_title),
                'accuracy': torch.mean(((logits > 0) == is_title).float()).detach()}

    def training_step(self, batch, batch_idx):
        return self.step(batch, batch_idx)

    def validation_step(self, batch, batch_idx):
        return self.step(batch, batch_idx)

    @staticmethod
    def calc_mean(outputs):

        result = defaultdict(lambda: 0)

        for output in outputs:
            for key, value in output.items():
                result[key] += value

        for key, value in result.items():
            result[key] = value / len(outputs)

        return result

    def training_epoch_end(self, outputs) -> None:
        print("---------------- Training Results ---------------")
        print(NamesClassifier.calc_mean(outputs))

    def validation_epoch_end(self, outputs) -> None:
        print("---------------- Validation Results ---------------")
        result = NamesClassifier.calc_mean(outputs)
        print(result)

        if result['accuracy'] > self.best_val_acc:
            torch.save(self.state_dict(), 'best_model.pth')

        torch.save(self.state_dict(), 'last_model.pth')

    def configure_optimizers(self):
        return torch.optim.Adam(params=self.parameters(), lr=1e-5)

    def classify(self, name):
        """
        :param name: A string.
        :return: A boolean that indicates whether the name is a title or not.
        """

        name = prepro(tokenizer=self.tokenizer, name=name)

        input_ids = torch.Tensor(name['input_ids']).long()
        attention_mask = torch.Tensor(name['attention_mask']).long()
        token_type_ids = torch.Tensor(name['token_type_ids']).long()

        with torch.no_grad():
            return bool(self(input_ids.unsqueeze(0), attention_mask.unsqueeze(0), token_type_ids.unsqueeze(0)).squeeze(0) > 0)

