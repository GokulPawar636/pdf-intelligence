from __future__ import annotations


SEMANTIC_EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "records": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "record_type": {
                        "type": "string"
                    },
                    "fields": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {
                                    "type": "string"
                                },
                                "raw_value": {
                                    "type": [
                                        "string",
                                        "null"
                                    ]
                                },
                                "normalized_value": {
                                    "type": [
                                        "string",
                                        "number",
                                        "boolean",
                                        "null"
                                    ]
                                },
                                "unit": {
                                    "type": [
                                        "string",
                                        "null"
                                    ]
                                },
                                "evidence_ids": {
                                    "type": "array",
                                    "items": {
                                        "type": "string"
                                    }
                                }
                            },
                            "required": [
                                "name",
                                "raw_value",
                                "normalized_value",
                                "unit",
                                "evidence_ids"
                            ],
                            "additionalProperties": False
                        }
                    }
                },
                "required": [
                    "record_type",
                    "fields"
                ],
                "additionalProperties": False
            }
        }
    },
    "required": [
        "records"
    ],
    "additionalProperties": False
}