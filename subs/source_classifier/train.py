
from dataset import ds
from model import SourceClassifier
import torch
import pytorch_lightning as pl
from torch.utils.data import DataLoader
from utils import organize_data, get_data
import random


if __name__ == "__main__":

    data, label_to_name = organize_data(get_data())
    validation_pct = 0.2

    train_data = []
    val_data = []

    for cls_sources, cls_label in data:
        random.shuffle(cls_sources)
        split_idx = int((1 - validation_pct) * len(cls_sources))
        split_idx = max(split_idx, 1)  # we take at list 1 element from each class for training
        cls_train = cls_sources[: split_idx]
        cls_val = cls_sources[split_idx:]

        train_data += [(cls_train, cls_label)]
        val_data += [(cls_val, cls_label)]

    train_dataset = ds(data=train_data, augment=True)
    validation_dataset = ds(data=val_data, augment=True)

    print(len(train_dataset))
    print(len(validation_dataset))

    model = SourceClassifier(num_cls=len(data), label_to_name=label_to_name)

    num_training_epochs = 1000000000000000

    train_dataloader = DataLoader(dataset=train_dataset,
                                  batch_size=64,
                                  shuffle=False,
                                  num_workers=16)

    validation_dataloader = DataLoader(dataset=validation_dataset,
                                       batch_size=64,
                                       shuffle=False,
                                       num_workers=16)

    trainer = pl.Trainer(max_epochs=num_training_epochs, gpus=1)
    trainer.fit(model, train_dataloader, validation_dataloader)
