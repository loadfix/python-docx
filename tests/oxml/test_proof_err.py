"""Unit-test suite for ``w:proofErr`` round-trip preservation."""

from __future__ import annotations

from typing import cast

from docx.oxml.ns import qn
from docx.oxml.shared import CT_ProofErr
from docx.oxml.text.paragraph import CT_P

from ..unitutil.cxml import element


class DescribeCT_ProofErr:
    """Unit-test suite for :class:`docx.oxml.shared.CT_ProofErr`."""

    def it_exposes_the_required_type_attribute(self):
        proof = cast(CT_ProofErr, element("w:proofErr{w:type=spellStart}"))

        assert proof.type == "spellStart"

    def it_round_trips_through_serialization(self):
        # -- matches the shape Word emits around mid-run proof markers --
        p = cast(
            CT_P,
            element(
                "w:p/("
                "w:proofErr{w:type=spellStart},"
                "w:r/w:t,"
                "w:proofErr{w:type=spellEnd})"
            ),
        )

        # -- after round-trip both markers are still present with their types --
        proofs = p.findall(qn("w:proofErr"))
        assert len(proofs) == 2
        assert proofs[0].get(qn("w:type")) == "spellStart"
        assert proofs[1].get(qn("w:type")) == "spellEnd"

    def it_is_registered_for_all_four_proof_types(self):
        for proof_type in ("spellStart", "spellEnd", "gramStart", "gramEnd"):
            proof = cast(
                CT_ProofErr, element("w:proofErr{w:type=%s}" % proof_type)
            )
            assert isinstance(proof, CT_ProofErr)
            assert proof.type == proof_type
