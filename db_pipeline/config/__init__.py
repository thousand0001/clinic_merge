"""診所與來源系統設定。"""

from .models import ClinicConfig, DetectionRule, load_clinic_config

__all__ = ["ClinicConfig", "DetectionRule", "load_clinic_config"]

