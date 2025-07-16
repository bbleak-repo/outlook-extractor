"""Export presets management for Outlook Extractor.

This module handles saving, loading, and managing export presets.
"""
import json
import logging
import uuid
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Any, List, Optional, TypedDict

logger = logging.getLogger(__name__)

class ExportPreset(TypedDict, total=False):
    """Type definition for export presets."""
    id: str
    name: str
    description: str
    format: str  # 'csv' or 'excel'
    include_headers: bool
    export_fields: List[str]
    options: Dict[str, Any]
    created_at: str
    updated_at: str

@dataclass
class ExportPresetManager:
    """Manages export presets with CRUD operations."""
    presets_dir: Path = field(default_factory=lambda: Path.home() / '.outlook_extractor' / 'presets')
    presets: Dict[str, ExportPreset] = field(default_factory=dict)
    
    def __post_init__(self):
        """Initialize the preset directory and load existing presets."""
        self.presets_dir.mkdir(parents=True, exist_ok=True)
        self._load_presets()
    
    def _load_presets(self) -> None:
        """Load all presets from the presets directory."""
        self.presets = {}
        for preset_file in self.presets_dir.glob('*.json'):
            try:
                with open(preset_file, 'r', encoding='utf-8') as f:
                    preset_data = json.load(f)
                    if self._validate_preset(preset_data):
                        self.presets[preset_data['id']] = preset_data
                    else:
                        logger.warning(f"Invalid preset format in {preset_file}")
            except Exception as e:
                logger.error(f"Error loading preset {preset_file}: {e}")
    
    def _validate_preset(self, preset_data: Dict[str, Any]) -> bool:
        """Validate a preset dictionary."""
        required_fields = ['id', 'name', 'format']
        return all(field in preset_data for field in required_fields)
    
    def _get_preset_path(self, preset_id: str) -> Path:
        """Get the filesystem path for a preset."""
        return self.presets_dir / f"{preset_id}.json"
    
    def create_preset(
        self,
        name: str,
        format: str,
        export_fields: List[str],
        include_headers: bool = True,
        description: str = "",
        options: Optional[Dict[str, Any]] = None,
        preset_id: Optional[str] = None
    ) -> ExportPreset:
        """Create a new export preset.
        
        Args:
            name: Name of the preset
            format: Export format ('csv' or 'excel')
            export_fields: List of field names to include
            include_headers: Whether to include headers in export
            description: Optional description
            options: Additional format-specific options
            preset_id: Optional preset ID (will generate if not provided)
            
        Returns:
            The created preset
        """
        now = datetime.now(timezone.utc).isoformat()
        preset_id = preset_id or str(uuid.uuid4())
        
        preset: ExportPreset = {
            'id': preset_id,
            'name': name,
            'description': description,
            'format': format,
            'export_fields': export_fields,
            'include_headers': include_headers,
            'options': options or {},
            'created_at': now,
            'updated_at': now
        }
        
        return self.save_preset(preset)
    
    def save_preset(self, preset: ExportPreset) -> ExportPreset:
        """Save a preset to disk."""
        if not self._validate_preset(preset):
            raise ValueError("Invalid preset format")
        
        # Update timestamps
        now = datetime.now(timezone.utc).isoformat()
        if 'created_at' not in preset:
            preset['created_at'] = now
        preset['updated_at'] = now
        
        # Save to file
        preset_path = self._get_preset_path(preset['id'])
        with open(preset_path, 'w', encoding='utf-8') as f:
            json.dump(preset, f, indent=2)
        
        # Update in-memory cache
        self.presets[preset['id']] = preset
        return preset
    
    def get_preset(self, preset_id: str) -> Optional[ExportPreset]:
        """Get a preset by ID."""
        return self.presets.get(preset_id)
    
    def get_all_presets(self) -> List[ExportPreset]:
        """Get all presets, sorted by name."""
        return sorted(
            list(self.presets.values()),
            key=lambda p: p['name'].lower()
        )
    
    def update_preset(self, preset_id: str, **updates) -> Optional[ExportPreset]:
        """Update a preset with new values."""
        if preset_id not in self.presets:
            return None
            
        preset = self.presets[preset_id].copy()
        preset.update(updates)
        return self.save_preset(preset)
    
    def delete_preset(self, preset_id: str) -> bool:
        """Delete a preset."""
        if preset_id not in self.presets:
            return False
            
        try:
            preset_path = self._get_preset_path(preset_id)
            if preset_path.exists():
                preset_path.unlink()
            del self.presets[preset_id]
            return True
        except Exception as e:
            logger.error(f"Error deleting preset {preset_id}: {e}")
            return False
    
    def get_default_presets(self) -> List[ExportPreset]:
        """Get default presets that are built into the application."""
        return [
            {
                'id': 'default_full_csv',
                'name': 'Full Export (CSV)',
                'description': 'Export all fields to CSV',
                'format': 'csv',
                'include_headers': True,
                'export_fields': [
                    'entry_id', 'conversation_id', 'subject', 'sender_name', 'sender_email',
                    'to_recipients', 'cc_recipients', 'bcc_recipients', 'received_time',
                    'sent_time', 'categories', 'importance', 'sensitivity', 'has_attachments',
                    'is_read', 'is_flagged', 'is_priority', 'is_admin', 'body', 'html_body',
                    'folder_path', 'thread_id', 'thread_depth'
                ],
                'options': {}
            },
            {
                'id': 'default_full_excel',
                'name': 'Full Export (Excel)',
                'description': 'Export all fields to Excel with formatting',
                'format': 'excel',
                'include_headers': True,
                'export_fields': [
                    'entry_id', 'conversation_id', 'subject', 'sender_name', 'sender_email',
                    'to_recipients', 'cc_recipients', 'bcc_recipients', 'received_time',
                    'sent_time', 'categories', 'importance', 'sensitivity', 'has_attachments',
                    'is_read', 'is_flagged', 'is_priority', 'is_admin', 'body', 'html_body',
                    'folder_path', 'thread_id', 'thread_depth'
                ],
                'options': {
                    'auto_size_columns': True,
                    'freeze_headers': True,
                    'add_filters': True
                }
            },
            {
                'id': 'default_metadata_only',
                'name': 'Metadata Only',
                'description': 'Export only metadata fields (no message bodies)', 
                'format': 'csv',
                'include_headers': True,
                'export_fields': [
                    'entry_id', 'conversation_id', 'subject', 'sender_name', 'sender_email',
                    'to_recipients', 'cc_recipients', 'bcc_recipients', 'received_time',
                    'sent_time', 'categories', 'importance', 'sensitivity', 'has_attachments',
                    'is_read', 'is_flagged', 'is_priority', 'is_admin', 'folder_path'
                ],
                'options': {}
            }
        ]
    
    def ensure_default_presets_exist(self) -> None:
        """Ensure default presets exist in the user's presets."""
        default_presets = {p['id']: p for p in self.get_default_presets()}
        
        # Add any missing default presets
        for preset_id, preset in default_presets.items():
            if preset_id not in self.presets:
                self.save_preset(preset.copy())
        
        # Update existing default presets if they've changed
        for preset_id in default_presets:
            if preset_id in self.presets:
                default_preset = default_presets[preset_id]
                current_preset = self.presets[preset_id]
                
                # Check if the preset needs updating
                needs_update = False
                for key, value in default_preset.items():
                    if key not in current_preset or current_preset[key] != value:
                        needs_update = True
                        break
                
                if needs_update:
                    # Preserve the created_at timestamp
                    created_at = current_preset.get('created_at')
                    updated_preset = default_preset.copy()
                    if created_at:
                        updated_preset['created_at'] = created_at
                    self.save_preset(updated_preset)
