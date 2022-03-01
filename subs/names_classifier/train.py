
from dataset import ds
from model import NamesClassifier
import torch
import pytorch_lightning as pl
from torch.utils.data import DataLoader


if __name__ == "__main__":

    model = NamesClassifier()

    ds = ds()
    train_size = int(len(ds) * 0.5)
    validation_size = len(ds) - train_size
    train_dataset, validation_dataset = torch.utils.data.random_split(ds, [train_size, validation_size])

    model = NamesClassifier()
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


