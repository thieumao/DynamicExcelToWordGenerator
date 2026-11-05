from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from docgen.domain.profile.io import load_profile
from docgen.domain.profile.models import Profile
from docgen.testing.construction import write_construction_kit
from docgen.testing.samples import write_customer_a, write_customer_b, write_customer_c


@dataclass
class Sample:
    profile_path: Path
    workbook: Path
    profile: Profile

    @property
    def directory(self) -> Path:
        return self.profile_path.parent


@pytest.fixture()
def sample_a(tmp_path: Path) -> Sample:
    profile_path, workbook = write_customer_a(tmp_path / "customer_a")
    return Sample(profile_path, workbook, load_profile(profile_path))


@pytest.fixture()
def sample_b(tmp_path: Path) -> Sample:
    profile_path, workbook = write_customer_b(tmp_path / "customer_b")
    return Sample(profile_path, workbook, load_profile(profile_path))


@pytest.fixture()
def sample_c(tmp_path: Path) -> Sample:
    profile_path, workbook = write_customer_c(tmp_path / "customer_c")
    return Sample(profile_path, workbook, load_profile(profile_path))


@pytest.fixture()
def construction(tmp_path: Path) -> Sample:
    profile_path, workbook = write_construction_kit(tmp_path / "xay_dung")
    return Sample(profile_path, workbook, load_profile(profile_path))
