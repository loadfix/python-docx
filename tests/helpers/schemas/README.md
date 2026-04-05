# OOXML Schema Files

This directory contains XSD schema files for validating OOXML XML parts.

## Bundled Schemas

- `wml-comments.xsd` — Simplified schema for `word/comments.xml` validation.

## Full ECMA-376 Schemas

For comprehensive schema validation, download the full XSD schemas from ECMA:

  https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

The relevant files are in Part 4 (Transitional Migration Features) of the standard.
Place the downloaded `.xsd` files in this directory and use `load_schema()` from
`tests/helpers/schema.py` to load them.

## How Bundled Schemas Work

The bundled schemas are simplified subsets of the full ECMA-376 schemas. They validate
the most important structural constraints for elements that python-docx produces, without
requiring the complete (very large) schema set. They use `processContents="lax"` for
child elements to allow content that goes beyond what the simplified schema defines.
