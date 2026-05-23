import os
import json
import tempfile
from pathlib import Path

import pytest
from src.modules.data_manager import DataManager


@pytest.fixture
def tmp_config():
    """每个测试用独立的临时JSON文件"""
    with tempfile.TemporaryDirectory() as d:
        yield Path(d) / "config.json"


def test_initial_empty(tmp_config):
    dm = DataManager(tmp_config)
    assert dm.get_total_count() == 0
    assert dm.get_name_list() == []
    assert not dm.has_name_list()


def test_update_and_read(tmp_config):
    dm = DataManager(tmp_config)
    dm.update_name_list(["张三", "李四", "王五"])
    assert dm.get_total_count() == 3
    assert dm.get_name_list() == ["张三", "李四", "王五"]
    assert dm.has_name_list()


def test_dedup_keeps_order(tmp_config):
    dm = DataManager(tmp_config)
    dm.update_name_list(["张三", "李四", "张三", "王五", "李四"])
    assert dm.get_name_list() == ["张三", "李四", "王五"]


def test_persistence(tmp_config):
    """写入后再创建一个新实例，应能读回数据"""
    dm1 = DataManager(tmp_config)
    dm1.update_name_list(["张三", "李四"])

    dm2 = DataManager(tmp_config)
    assert dm2.get_name_list() == ["张三", "李四"]


def test_corrupted_json_fallback(tmp_config):
    """JSON损坏时应能回退到空结构而不崩溃"""
    tmp_config.parent.mkdir(parents=True, exist_ok=True)
    tmp_config.write_text("{not valid json", encoding="utf-8")
    dm = DataManager(tmp_config)
    assert dm.get_total_count() == 0


def test_senior_config_default_values(tmp_config):
    """新增配置字段默认值"""
    dm = DataManager(tmp_config)
    assert dm.get_senior_assistants() == []
    assert dm.is_senior_should_fixed_enabled() is False
    assert dm.get_name_grades() == {}


def test_set_senior_assistants_filters_invalid_and_dedup(tmp_config):
    """大四助理名单：仅保留总名单内成员，且去重保序"""
    dm = DataManager(tmp_config)
    dm.update_name_list(["张三", "李四", "王五"])

    dm.set_senior_assistants(["张三", "李四", "张三", "不存在"])
    assert dm.get_senior_assistants() == ["张三", "李四"]


def test_toggle_senior_rule_persistence(tmp_config):
    """大四规则开关应可持久化"""
    dm = DataManager(tmp_config)
    dm.set_senior_should_fixed_enabled(True)
    assert dm.is_senior_should_fixed_enabled() is True

    dm2 = DataManager(tmp_config)
    assert dm2.is_senior_should_fixed_enabled() is True


def test_update_name_list_cleans_removed_seniors(tmp_config):
    """更新总名单时，自动移除已不存在的大四助理"""
    dm = DataManager(tmp_config)
    dm.update_name_list(["张三", "李四", "王五"])
    dm.set_senior_assistants(["张三", "李四"])

    dm.update_name_list(["张三", "王五"])
    assert dm.get_senior_assistants() == ["张三"]


def test_add_name_with_grade(tmp_config):
    """手动添加姓名时保存年级并追加到总名单末尾"""
    dm = DataManager(tmp_config)

    assert dm.add_name("张三", "大二") is True
    assert dm.get_name_list() == ["张三"]
    assert dm.get_total_count() == 1
    assert dm.get_name_grade("张三") == "大二"


def test_add_senior_name_adds_to_senior_assistants(tmp_config):
    """手动添加大四姓名时自动计入大四助理名单"""
    dm = DataManager(tmp_config)

    assert dm.add_name("李四", "大四") is True
    assert dm.get_name_list() == ["李四"]
    assert dm.get_senior_assistants() == ["李四"]


def test_add_duplicate_name_rejected(tmp_config):
    """重复添加姓名不会覆盖原有名单和年级"""
    dm = DataManager(tmp_config)
    dm.add_name("张三", "大一")

    assert dm.add_name("张三", "大四") is False
    assert dm.get_name_list() == ["张三"]
    assert dm.get_name_grade("张三") == "大一"
    assert dm.get_senior_assistants() == []


def test_update_name_list_cleans_removed_grades(tmp_config):
    """更新总名单时，自动移除已不存在人员的年级信息"""
    dm = DataManager(tmp_config)
    dm.add_name("张三", "大二")
    dm.add_name("李四", "大三")

    dm.update_name_list(["张三", "王五"])
    assert dm.get_name_grades() == {"张三": "大二"}
