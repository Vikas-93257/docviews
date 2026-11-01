#!/usr/bin/env python3
"""Compare the documented python-docx 1.2.0 API inventory with docviews.

This first-pass audit measures import/re-export paths, sample-object member
availability, callable parameter shapes and annotations, property writability,
and enum values. It does not treat access as proof of behavior: only records
listed in an explicit differential-test manifest receive a tested behavior status.

Example:

    python tools/compare_api_parity.py \
        --reference research/python-docx-v1.2.0 \
        --inventory docs/api-parity/python_docx_1_2_0_inventory.json \
        --output docs/api-parity/python_docx_1_2_0_comparison.json

Run with the maintainer/test environment that has python-docx's lxml dependency
and pytest installed. The upstream source path is inserted ahead of site-packages
so the comparison is pinned to the recorded 1.2.0 source snapshot.
"""

from __future__ import annotations

import argparse
import importlib
import inspect
import json
import sys
from collections import Counter
from datetime import date, datetime
from enum import Enum
from pathlib import Path
from typing import Any

UPSTREAM_COMMIT = "e45454602b53e8e572b179ccf1c91093ec9f4ed7"

# The API docs name these objects as classes. The compat bridge currently
# supplies equivalent proxy instances through these representative recipes.
OWNER_SAMPLE = {
    "docx.comments.Comment": "Comment",
    "docx.comments.Comments": "Comments",
    "docx.dml.color.ColorFormat": "ColorFormat",
    "docx.document.Document": "Document",
    "docx.opc.coreprops.CoreProperties": "CoreProperties",
    "docx.section.Section": "Section",
    "docx.section.Sections": "Sections",
    "docx.section._Footer": "Footer",
    "docx.section._Header": "Header",
    "docx.settings.Settings": "Settings",
    "docx.shape.InlineShape": "InlineShape",
    "docx.shape.InlineShapes": "InlineShapes",
    "docx.shared.Length": "Length",
    "docx.shared.RGBColor": "RGBColor",
    "docx.styles.latent.LatentStyles": "LatentStyles",
    "docx.styles.latent._LatentStyle": "LatentStyle",
    "docx.styles.style.BaseStyle": "BaseStyle",
    "docx.styles.style.CharacterStyle": "CharacterStyle",
    "docx.styles.style.ParagraphStyle": "ParagraphStyle",
    "docx.styles.style._NumberingStyle": "NumberingStyle",
    "docx.styles.style._TableStyle": "TableStyle",
    "docx.styles.styles.Styles": "Styles",
    "docx.table.Table": "Table",
    "docx.table._Cell": "Cell",
    "docx.table._Column": "Column",
    "docx.table._Columns": "Columns",
    "docx.table._Row": "Row",
    "docx.table._Rows": "Rows",
    "docx.text.font.Font": "Font",
    "docx.text.hyperlink.Hyperlink": "Hyperlink",
    "docx.text.pagebreak.RenderedPageBreak": "RenderedPageBreak",
    "docx.text.paragraph.Paragraph": "Paragraph",
    "docx.text.parfmt.ParagraphFormat": "ParagraphFormat",
    "docx.text.run.Font": "Font",
    "docx.text.run.Run": "Run",
    "docx.text.tabstops.TabStop": "TabStop",
    "docx.text.tabstops.TabStops": "TabStops",
}

MIRRORED_CLASS_MODULES = frozenset({
    "docx.dml.color", "docx.opc.coreprops", "docx.section", "docx.settings",
    "docx.shape", "docx.styles.latent", "docx.styles.style",
    "docx.styles.styles", "docx.table", "docx.text.font",
    "docx.text.hyperlink", "docx.text.pagebreak", "docx.text.paragraph",
    "docx.text.parfmt", "docx.text.run", "docx.text.tabstops",
})

STRUCTURAL = {"part", "element", "blob"}


def add_style_samples(reference_samples: dict, shim_samples: dict, reference_doc,
                      shim_doc, reference_docx, shim) -> None:
    """Add style subtypes not represented by the original bounded sample."""
    reference_styles = reference_doc.styles
    shim_styles = shim_doc.styles
    reference_samples["BaseStyle"] = reference_styles["Normal"]
    shim_samples["BaseStyle"] = shim_styles["Normal"]
    reference_samples["ParagraphStyle"] = reference_styles["Heading 1"]
    shim_samples["ParagraphStyle"] = shim_styles["Heading 1"]
    reference_samples["CharacterStyle"] = reference_styles["Default Paragraph Font"]
    shim_samples["CharacterStyle"] = shim_styles["Default Paragraph Font"]
    reference_samples["TableStyle"] = reference_styles["Table Grid"]
    shim_samples["TableStyle"] = shim_styles["Table Grid"]
    # LIST is a documented WD_STYLE_TYPE. Keep the actual returned proxies so
    # the audit can reveal if the replacement changes the type or members.
    ref_list_type = reference_docx.enum.style.WD_STYLE_TYPE.LIST
    shim_list_type = shim.enum.WD_STYLE_TYPE.LIST
    reference_samples["NumberingStyle"] = reference_styles.add_style(
        "API parity inventory numbering style", ref_list_type
    )
    shim_samples["NumberingStyle"] = shim_styles.add_style(
        "API parity inventory numbering style", shim_list_type
    )


def safe_repr(value: Any) -> Any:
    if value is inspect.Signature.empty:
        return None
    if isinstance(value, Enum):
        return {
            "enum_type": type(value).__name__,
            "name": value.name,
            "value": safe_repr(value.value),
        }
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, tuple):
        return [safe_repr(item) for item in value]
    return repr(value)


def annotation_text(value: Any) -> str | None:
    if value is inspect.Signature.empty:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, type):
        return f"{value.__module__}.{value.__qualname__}"
    return str(value)


def signature_data(callable_obj: Any) -> dict[str, Any] | None:
    try:
        signature = inspect.signature(callable_obj)
    except (TypeError, ValueError):
        return None
    parameters = []
    shape = []
    annotations = []
    for parameter in signature.parameters.values():
        default_missing = parameter.default is inspect.Parameter.empty
        default = safe_repr(parameter.default)
        parameters.append({
            "name": parameter.name,
            "kind": parameter.kind.name,
            "required": default_missing,
            "default": default,
            "annotation": annotation_text(parameter.annotation),
        })
        shape.append((parameter.name, parameter.kind.name, default_missing,
                      json.dumps(default, sort_keys=True, ensure_ascii=False)))
        annotations.append(annotation_text(parameter.annotation))
    return {
        "display": str(signature),
        "parameters": parameters,
        "shape": shape,
        "annotations": annotations,
        "return_annotation": annotation_text(signature.return_annotation),
    }


def annotation_signature(data: dict[str, Any]) -> tuple[Any, ...]:
    return tuple(data["annotations"]), data["return_annotation"]


def static_descriptor(obj: Any, name: str) -> Any:
    try:
        return inspect.getattr_static(type(obj), name)
    except AttributeError:
        return None


def is_writable_descriptor(descriptor: Any) -> bool | None:
    if descriptor is None:
        return None
    if isinstance(descriptor, property):
        return descriptor.fset is not None
    # python-docx's lazyproperty implements __set__ only to cache a computed
    # value; it is not a documented setter. Treat it as read-only here.
    if type(descriptor).__name__ == "lazyproperty":
        return False
    if hasattr(descriptor, "__set__"):
        return True
    return False


def get_member(obj: Any, name: str) -> tuple[bool, Any, str | None]:
    try:
        return True, getattr(obj, name), None
    except AttributeError as exc:
        descriptor = static_descriptor(obj, name)
        if descriptor is not None:
            return True, None, f"getter raised AttributeError: {exc}"
        return False, None, str(exc)
    except Exception as exc:  # a descriptor exists but its read raised
        descriptor = static_descriptor(obj, name)
        if descriptor is not None:
            return True, None, f"getter raised {type(exc).__name__}: {exc}"
        return False, None, f"{type(exc).__name__}: {exc}"


def target_import_path(qualified_name: str, kind: str) -> tuple[str, str]:
    """Return migration-path classification and best documented target path."""
    if qualified_name == "docx.Document":
        return "reexported_top_level", "docviews.compat.Document"
    if kind == "class":
        module, _, member = qualified_name.rpartition(".")
        if module in MIRRORED_CLASS_MODULES:
            target = "docviews.compat." + module.removeprefix("docx.")
            return "mirrored_module", f"{target}.{member}"
    if qualified_name.startswith("docx.shared."):
        return "mirrored_module", "docviews.compat.shared." + qualified_name.rsplit(".", 1)[-1]
    if qualified_name.startswith("docx.comments."):
        return "mirrored_module", "docviews.compat.comments." + qualified_name.rsplit(".", 1)[-1]
    if qualified_name.startswith("docx.enum."):
        return "mirrored_module", "docviews.compat.enum." + qualified_name[len("docx.enum."):]
    if qualified_name == "docx.document.Document":
        return "top_level_reexport_only", "docviews.compat.Document"
    if qualified_name.startswith("docx."):
        return "object_access_only", "returned through docviews.compat.Document proxies"
    return "not_mapped", ""


def import_candidate(path: str) -> Any:
    """Resolve a ``docviews.compat`` import path without evaluating expressions."""
    if not path or "returned through" in path:
        return None
    parts = path.split(".")
    for split_at in range(len(parts), 0, -1):
        module_name = ".".join(parts[:split_at])
        try:
            value: Any = importlib.import_module(module_name)
        except Exception:
            continue
        for component in parts[split_at:]:
            try:
                value = getattr(value, component)
            except Exception:
                return None
        return value
    return None


def compare_signatures(reference_callable: Any, candidate_callable: Any) -> dict[str, Any]:
    reference = signature_data(reference_callable)
    candidate = signature_data(candidate_callable)
    if reference is None or candidate is None:
        return {
            "signature_parity": "signature_unavailable",
            "reference_signature_live": reference["display"] if reference else None,
            "docviews_signature_live": candidate["display"] if candidate else None,
            "reference_call_shape": reference["shape"] if reference else None,
            "docviews_call_shape": candidate["shape"] if candidate else None,
            "annotation_parity": "not_compared",
        }
    shape_match = reference["shape"] == candidate["shape"]
    annotations_match = annotation_signature(reference) == annotation_signature(candidate)
    return {
        "signature_parity": "call_shape_match" if shape_match else "call_shape_mismatch",
        "reference_signature_live": reference["display"],
        "docviews_signature_live": candidate["display"],
        "reference_call_shape": reference["shape"],
        "docviews_call_shape": candidate["shape"],
        "annotation_parity": "match" if annotations_match else "mismatch",
        "reference_annotations": annotation_signature(reference),
        "docviews_annotations": annotation_signature(candidate),
    }


def api_owner_status(record: dict[str, Any], reference_samples: dict,
                     shim_samples: dict) -> dict[str, Any]:
    name = record["qualified_name"]
    if record["kind"] in {"class", "function"}:
        return compare_api_class(record, reference_samples, shim_samples)
    owner, member = name.rsplit(".", 1)
    label = OWNER_SAMPLE.get(owner)
    if label is None or label not in reference_samples or label not in shim_samples:
        return {
            "availability": "sample_mapping_missing",
            "signature_parity": "not_assessed",
            "annotation_parity": "not_assessed",
            "behavior_parity": "not_assessed",
            "owner_sample": label,
            "audit_note": "No representative instance was mapped for this documented proxy class.",
        }
    reference_obj = reference_samples[label]
    shim_obj = shim_samples[label]
    ref_present, ref_value, ref_error = get_member(reference_obj, member)
    shim_present, shim_value, shim_error = get_member(shim_obj, member)
    availability = (
        "available" if shim_present else "missing"
    ) if ref_present else "reference_sample_missing"
    result: dict[str, Any] = {
        "availability": availability,
        "owner_sample": label,
        "reference_member_access_error": ref_error,
        "docviews_member_access_error": shim_error,
        "behavior_parity": "not_assessed",
    }
    if record["kind"] == "method":
        if not ref_present or not callable(ref_value) or not shim_present or not callable(shim_value):
            result.update({
                "signature_parity": "not_callable_or_missing",
                "annotation_parity": "not_compared",
            })
        else:
            result.update(compare_signatures(ref_value, shim_value))
    else:
        result["signature_parity"] = "not_applicable_property_or_attribute"
        result["annotation_parity"] = "not_compared"
        ref_write = is_writable_descriptor(static_descriptor(reference_obj, member))
        shim_write = is_writable_descriptor(static_descriptor(shim_obj, member))
        result["reference_writable"] = ref_write
        result["docviews_writable"] = shim_write
        if ref_write is None or shim_write is None:
            result["writability_parity"] = "not_determined"
        else:
            result["writability_parity"] = "match" if ref_write == shim_write else "mismatch"
    result["import_path_status"], result["compat_import_path"] = target_import_path(name, record["kind"])
    return result


def compare_api_class(record: dict[str, Any], reference_samples: dict,
                      shim_samples: dict) -> dict[str, Any]:
    name = record["qualified_name"]
    import_status, target_path = target_import_path(name, record["kind"])
    candidate = import_candidate(target_path)
    sample_label = OWNER_SAMPLE.get(name)
    reference_sample = reference_samples.get(sample_label) if sample_label else None
    shim_sample = shim_samples.get(sample_label) if sample_label else None
    if candidate is not None:
        availability = "available"
    elif shim_sample is not None:
        availability = "available_as_proxy_instance"
    else:
        availability = "not_directly_importable"
    result: dict[str, Any] = {
        "import_path_status": import_status,
        "compat_import_path": target_path,
        "availability": availability,
        "owner_sample": sample_label,
        "reference_runtime_type": (
            f"{type(reference_sample).__module__}.{type(reference_sample).__name__}"
            if reference_sample is not None else None
        ),
        "docviews_runtime_type": (
            f"{type(shim_sample).__module__}.{type(shim_sample).__name__}"
            if shim_sample is not None else (
                f"{candidate.__module__}.{candidate.__name__}"
                if isinstance(candidate, type) else None
            )
        ),
        "behavior_parity": "not_assessed",
    }
    if reference_sample is not None and shim_sample is not None:
        ref_name = type(reference_sample).__name__
        shim_name = type(shim_sample).__name__
        result["proxy_type_parity"] = "same_name" if ref_name == shim_name else "different_proxy_implementation"
    if record["kind"] == "function":
        module_name, object_name = name.rsplit(".", 1)
        reference = getattr(importlib.import_module(module_name), object_name)
        if candidate is not None:
            result.update(compare_signatures(reference, candidate))
        else:
            result["signature_parity"] = "candidate_not_importable"
            result["annotation_parity"] = "not_compared"
        return result

    # Direct constructors for shared value objects are part of the user API.
    # The top-level Document factory is checked separately as a function.
    if name.startswith("docx.shared.") and candidate is not None:
        module_name, object_name = name.rsplit(".", 1)
        reference = getattr(importlib.import_module(module_name), object_name)
        result.update(compare_signatures(reference, candidate))
        result["constructor_scope"] = "in_scope_value_constructor"
        return result

    signature = record.get("signature", "")
    if any(token in signature for token in ("CT_", "DocumentPart", "TableParent", "ProvidesStoryPart")):
        result["signature_parity"] = "excluded_internal_constructor"
        result["annotation_parity"] = "excluded_internal_constructor"
        result["constructor_scope"] = "excluded_internal_constructor"
    else:
        result["signature_parity"] = "not_a_documented_public_constructor"
        result["annotation_parity"] = "not_applicable"
        result["constructor_scope"] = "proxy_or_factory_object"
    return result


def enum_module_for(record: dict[str, Any]) -> tuple[str, str] | None:
    path = record.get("upstream_runtime_name") or record["qualified_name"]
    if not path.startswith("docx.enum."):
        return None
    parts = path.split(".")
    if len(parts) < 4:
        return None
    module_name = ".".join(parts[:3])
    class_name = parts[3]
    return module_name, class_name


def compare_enum_record(record: dict[str, Any]) -> dict[str, Any]:
    name = record["qualified_name"]
    import_status, target_path = target_import_path(name, record["kind"])
    result: dict[str, Any] = {
        "import_path_status": import_status,
        "compat_import_path": target_path,
        "behavior_parity": "not_assessed",
    }
    if record["kind"] == "enum_alias":
        alias_path = name
        alias_module, alias_name = alias_path.rsplit(".", 1)
        reference_alias = getattr(importlib.import_module(alias_module), alias_name, None)
        candidate_alias = import_candidate(target_path)
        canonical = record.get("alias_of", "")
        canonical_module, canonical_name = canonical.rsplit(".", 1)
        reference_canonical = getattr(importlib.import_module(canonical_module), canonical_name, None)
        candidate_canonical = import_candidate(
            target_path.rsplit(".", 1)[0] + "." + canonical_name
        )
        result.update({
            "availability": "available" if candidate_alias is not None else "missing",
            "reference_alias_identity": reference_alias is reference_canonical,
            "docviews_alias_identity": candidate_alias is candidate_canonical,
            "alias_identity_parity": (
                "match" if (reference_alias is reference_canonical) ==
                (candidate_alias is candidate_canonical) else "mismatch"
            ),
            "signature_parity": "not_applicable",
        })
        return result

    module_path = "docx.enum." + name.split(".")[2]
    class_name = name.split(".")[3]
    ref_module = importlib.import_module(module_path)
    shim_module = importlib.import_module("docviews.compat.enum." + name.split(".")[2])
    ref_cls = getattr(ref_module, class_name, None)
    shim_cls = getattr(shim_module, class_name, None)
    if record["kind"] == "enum_class":
        reference_members = set(getattr(ref_cls, "__members__", {})) if ref_cls else set()
        shim_members = set(getattr(shim_cls, "__members__", {})) if shim_cls else set()
        result.update({
            "availability": "available" if shim_cls is not None else "missing",
            "signature_parity": "not_applicable_enum_type",
            "reference_member_count": len(reference_members),
            "docviews_member_count": len(shim_members),
            "enum_member_set_parity": "match" if reference_members == shim_members else "mismatch",
            "reference_members_missing_in_docviews": sorted(reference_members - shim_members),
            "docviews_extra_members": sorted(shim_members - reference_members),
        })
        return result

    member_name = name.rsplit(".", 1)[-1]
    ref_member = getattr(ref_cls, member_name, None) if ref_cls is not None else None
    shim_member = getattr(shim_cls, member_name, None) if shim_cls is not None else None
    upstream_exists = bool(record.get("present_in_python_docx_1_2_0"))
    if not upstream_exists:
        documented_value = record.get("documented_value")
        shim_value = getattr(shim_member, "value", None) if shim_member is not None else None
        shim_xml = getattr(shim_member, "xml_value", None) if shim_member is not None else None
        result.update({
            "availability": "documented_but_missing_upstream_runtime",
            "signature_parity": "not_applicable_enum_member",
            "value_parity": "reference_documentation_runtime_conflict",
            "documented_value": safe_repr(documented_value),
            "docviews_member_available": shim_member is not None,
            "docviews_value": safe_repr(shim_value),
            "documented_value_parity": (
                "match" if shim_member is not None and documented_value is not None
                and shim_value == documented_value else "not_compared"
            ),
            "xml_value_parity": "not_applicable",
            "docviews_xml_value": shim_xml,
            "docviews_str": str(shim_member) if shim_member is not None else None,
            "docviews_repr": repr(shim_member) if shim_member is not None else None,
            "reference_note": record.get("reference_note", ""),
            "documented_value_source": record.get("documented_value_source"),
        })
        return result
    if ref_member is None:
        result.update({
            "availability": "reference_member_missing",
            "signature_parity": "not_applicable_enum_member",
            "value_parity": "reference_member_missing",
            "xml_value_parity": "not_applicable",
        })
        return result
    available = shim_member is not None
    ref_value = getattr(ref_member, "value", None)
    shim_value = getattr(shim_member, "value", None) if available else None
    ref_xml = getattr(ref_member, "xml_value", None)
    shim_xml = getattr(shim_member, "xml_value", None) if available else None
    ref_str = str(ref_member)
    shim_str = str(shim_member) if available else None
    ref_repr = repr(ref_member)
    shim_repr = repr(shim_member) if available else None
    value_equal = available and ref_value == shim_value
    xml_equal = available and ref_xml == shim_xml
    xml_status = "match" if xml_equal else ("missing" if not available else "mismatch")
    if available and ref_xml is None and shim_xml is not None:
        xml_status = "docviews_extra_xml_token"
    result.update({
        "availability": "available" if available else "missing",
        "signature_parity": "not_applicable_enum_member",
        "reference_value": safe_repr(ref_value),
        "docviews_value": safe_repr(shim_value),
        "value_parity": "match" if value_equal else ("missing" if not available else "mismatch"),
        "reference_xml_value": ref_xml,
        "docviews_xml_value": shim_xml,
        "xml_value_parity": xml_status,
        "reference_str": ref_str,
        "docviews_str": shim_str,
        "string_parity": "match" if available and ref_str == shim_str else "mismatch",
        "reference_repr": ref_repr,
        "docviews_repr": shim_repr,
        "repr_parity": "match" if available and ref_repr == shim_repr else "mismatch",
    })
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", type=Path, required=True,
                        help="python-docx v1.2.0 source snapshot")
    parser.add_argument("--inventory", type=Path, default=Path(
        "docs/api-parity/python_docx_1_2_0_inventory.json"))
    parser.add_argument("--output", type=Path, default=Path(
        "docs/api-parity/python_docx_1_2_0_comparison.json"))
    parser.add_argument("--behavior", type=Path, default=Path(
        "docs/api-parity/python_docx_1_2_0_behavior.json"),
        help="optional manifest of behavior statuses backed by differential tests")
    args = parser.parse_args()
    project = Path(__file__).resolve().parents[1]
    reference = args.reference.resolve()
    sys.path.insert(0, str(reference / "src"))
    sys.path.insert(0, str(project / "src"))
    sys.path.insert(0, str(project / "tests"))
    sys.path.insert(0, str(project))

    import docx
    import docviews.compat as shim
    import compat_coverage

    reference_samples = compat_coverage.build(docx, compat_coverage.REAL_STYLE_TYPE)
    shim_samples = compat_coverage.build(shim, compat_coverage.OUR_STYLE_TYPE)
    add_style_samples(
        reference_samples,
        shim_samples,
        reference_samples["Document"],
        shim_samples["Document"],
        docx,
        shim,
    )
    inventory = json.loads(args.inventory.read_text(encoding="utf-8"))
    results: dict[str, dict[str, Any]] = {}
    for record in inventory["records"]:
        name = record["qualified_name"]
        if record.get("scope") == "excluded_internals":
            result = {
                "availability": "excluded_by_scope",
                "signature_parity": "excluded_by_scope",
                "behavior_parity": "excluded_by_scope",
                "import_path_status": "excluded_by_scope",
            }
        elif record["kind"].startswith("enum_"):
            result = compare_enum_record(record)
        else:
            result = api_owner_status(record, reference_samples, shim_samples)
        results[name] = result

    behavior_manifest = {}
    if args.behavior is not None and args.behavior.is_file():
        behavior_manifest = json.loads(args.behavior.read_text(encoding="utf-8"))
    allowed_behavior_statuses = {
        "tested_match", "tested_partial_match", "tested_mismatch",
        "intentional_difference",
    }
    for name, assessment in behavior_manifest.get("assessments", {}).items():
        if name not in results:
            print(f"WARNING: behavior assessment is not in inventory: {name}")
            continue
        status = assessment.get("status")
        if status not in allowed_behavior_statuses:
            raise ValueError(f"invalid behavior status {status!r} for {name}")
        results[name]["behavior_parity"] = status
        results[name]["behavior_evidence"] = assessment.get("evidence", [])
        if assessment.get("notes"):
            results[name]["behavior_notes"] = assessment["notes"]
        if "writability_parity" in assessment:
            writability = assessment["writability_parity"]
            if writability not in {"match", "mismatch", "not_determined"}:
                raise ValueError(
                    f"invalid writability status {writability!r} for {name}")
            results[name]["writability_parity"] = writability
            for key in ("reference_writable", "docviews_writable"):
                if key in assessment:
                    results[name][key] = assessment[key]

    status_counts = Counter(result.get("availability", "unknown") for result in results.values())
    call_shape_counts = Counter(
        result.get("signature_parity", "not_assessed")
        for result in results.values()
    )
    annotation_counts = Counter(
        result.get("annotation_parity", "not_compared")
        for result in results.values()
    )
    import_counts = Counter(
        result.get("import_path_status", "unknown")
        for result in results.values()
    )
    behavior_counts = Counter(
        result.get("behavior_parity", "not_assessed")
        for result in results.values()
    )
    writability_counts = Counter(
        result.get("writability_parity", "not_applicable")
        for result in results.values()
        if "writability_parity" in result
    )
    output = {
        "reference": {
            "package": "python-docx",
            "version": "1.2.0",
            "commit": UPSTREAM_COMMIT,
        },
        "candidate": {
            "package": "docviews.compat",
            "project_root": ".",
        },
        "comparison_date": date.today().isoformat(),
        "method": {
            "availability": "representative synthetic instances; no user documents",
            "signatures": "Python call shape (parameter names, kinds, required/default) plus annotations; methods are bound on samples",
            "behavior": "only explicit manifest entries backed by differential tests receive a tested status; other rows remain not_assessed",
            "writability": "static descriptors plus differential setter/read-only probes explicitly recorded in the behavior manifest",
            "behavior_manifest": str(args.behavior) if behavior_manifest else None,
            "scope": "documented API rows only; structural lxml/OOXML/OPC internals excluded",
        },
        "user_guide_recipes": behavior_manifest.get("user_guide_recipes", []),
        "counts": {
            "records": len(results),
            "availability": dict(sorted(status_counts.items())),
            "signature_parity": dict(sorted(call_shape_counts.items())),
            "annotation_parity": dict(sorted(annotation_counts.items())),
            "import_path_status": dict(sorted(import_counts.items())),
            "behavior_parity": dict(sorted(behavior_counts.items())),
            "writability_parity": dict(sorted(writability_counts.items())),
        },
        "records": results,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, indent=2, ensure_ascii=False) + "\n",
                           encoding="utf-8")
    print(f"Wrote comparisons for {len(results)} inventory entries to {args.output}")
    print("Availability:", dict(sorted(status_counts.items())))
    print("Signature:", dict(sorted(call_shape_counts.items())))
    print("Annotations:", dict(sorted(annotation_counts.items())))
    print("Import paths:", dict(sorted(import_counts.items())))
    print("Behavior evidence:", dict(sorted(behavior_counts.items())))
    print("Writability:", dict(sorted(writability_counts.items())))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
