import sys
from dataclasses import dataclass
from mapping import sells_id, sells_left, sells_price


@dataclass
class Sells:
    title: []
    left: []
    price: []


@dataclass
class ReceiptLine:
    title: []
    amount: []
    price: []
    bulk_price: []
