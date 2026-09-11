import json
from types import SimpleNamespace
import pytest
from app.semantic import excel_mapping


def fake_client(monkeypatch, proposal):
    monkeypatch.setenv('GROQ_API_KEY','synthetic-test-key')
    response = SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(proposal)))])
    client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=lambda **kwargs:response)),close=lambda:None)
    monkeypatch.setattr(excel_mapping,'Groq',lambda **kwargs:client)


def test_ai_cannot_map_one_column_to_nominal_and_reading(monkeypatch):
    fake_client(monkeypatch,{'header_row':0,'columns':{'object':0,'control':1,'nominal':2,
                 'tolerance':None,'unit':None,'lower':None,'upper':None},'readings':[2]})
    with pytest.raises(ValueError,match='overlapping'):
        excel_mapping.suggest_mapping([['Feature','Type','Target']])


def test_ai_mapping_validates_source_indices(monkeypatch):
    fake_client(monkeypatch,{'header_row':0,'columns':{'object':0,'control':1,'nominal':None,
                 'tolerance':None,'unit':None,'lower':None,'upper':None},'readings':[7]})
    with pytest.raises(ValueError,match='nonexistent'):
        excel_mapping.suggest_mapping([['Feature','Type','Observation']])


def test_ai_reports_unsupported_layout(monkeypatch):
    fake_client(monkeypatch,{'unsupported':True})
    assert excel_mapping.suggest_mapping([['Invoice','Amount']]) is None
