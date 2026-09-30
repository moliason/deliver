import requests
import httpx
import re
import os
CLIENT_ID = ""
CLIENT_SECRET = ""
UNION_ID = ""
SPACE_ID = ""
folder_ids = []
series_paths = {}
_api = None


def configure_dingtalk(config):
    global CLIENT_ID, CLIENT_SECRET, UNION_ID, SPACE_ID, folder_ids, series_paths, _api
    CLIENT_ID = os.getenv("DINGTALK_CLIENT_ID", "")
    CLIENT_SECRET = os.getenv("DINGTALK_CLIENT_SECRET", "")
    UNION_ID = os.getenv("DINGTALK_UNION_ID", "")
    SPACE_ID = str(config.get("space_id", ""))
    folder_ids = config.get("folder_ids", [])
    series_paths = {key.casefold(): value for key, value in config.get("series_paths", {}).items()}
    if config.get("enabled"):
        if not all((CLIENT_ID, CLIENT_SECRET, UNION_ID, SPACE_ID, folder_ids)):
            raise ValueError("启用钉钉需要三个 DINGTALK_* 环境变量、space_id 和 folder_ids")
        _api = httpx.Client(base_url="https://api.dingtalk.com", timeout=30)


SERIES_REQUIREMENT_KEYWORDS = ("未实现备注", "QA", "Xbuilder")

change_application_form = {}

def series_requirement_keyword(value):
    normalized = re.sub(r"[\W_]+", "", str(value or ""), flags=re.UNICODE).casefold()
    for keyword in SERIES_REQUIREMENT_KEYWORDS:
        if keyword.casefold() in normalized:
            return keyword
    return ""

#比对版本号.后整数对比
def _version_key(version):
    return tuple(int(part) for part in version.split("."))

def change_application_version_error(deliverable, forms):
    expected_match = re.fullmatch(r"变更申请单_(v.+)", str(deliverable).strip())
    expected = expected_match.group(1) if expected_match else ""
    actual_versions = {
        f"v{str(form.get('version', '')).strip()}"
        for form in forms.values()
        if str(form.get("version", "")).strip()
    }
    if expected and expected in actual_versions:
        return ""
    actual = ",".join(sorted(actual_versions))
    return f"变更申请单版本不一致：Redmine={expected or '<缺少版本>'}，实际拉取={actual or '<缺少版本>'}"

def get_access_token():
    if _api is None:
        raise ValueError("交付项需要钉钉文件，请先配置并启用 dingtalk")
    response = _api.post("/v1.0/oauth2/accessToken", json={"appKey": CLIENT_ID, "appSecret": CLIENT_SECRET})
    response.raise_for_status()
    return response.json()["accessToken"]

def get_all_files(accessToken, parent_id="0", series="", target_mode="", loop=False, target_keyword="变更申请", collect_all_change_forms=False):
    target_requirement = series_requirement_keyword(target_keyword)
    uses_analysis_version = target_keyword in ("外设功能分析表", "引脚功能缺失清单") or bool(target_requirement)
    response = _api.get(f"/v1.0/storage/spaces/{SPACE_ID}/dentries",
                         headers={"x-acs-dingtalk-access-token":accessToken},
                         params={"parentId": parent_id, "unionId": UNION_ID})
    response.raise_for_status()
    dentries = response.json()['dentries']
    if target_keyword == "变更申请" and any(
        item["type"] == "FOLDER"
        and os.path.basename(os.path.dirname(item["path"])).casefold() == target_mode.casefold()
        for item in dentries
    ):
        collect_all_change_forms = True

    if len(change_application_form) > 0 and not target_requirement and not collect_all_change_forms:
        return

    for item in dentries:
        if loop or series.casefold() == item["name"].casefold() or any(
            item["path"] == prefix or item["path"].startswith(prefix.rstrip("/") + "/")
            for prefix in series_paths.get(series.casefold(), [])
        ):
            if item["type"] == "FOLDER" and (len(change_application_form) == 0 or target_requirement or collect_all_change_forms):
                get_all_files(accessToken, item["id"], series, target_mode, True, target_keyword, collect_all_change_forms)

                if len(change_application_form) > 0 and not target_requirement and not collect_all_change_forms:
                    return
            elif item["type"] == "FILE" and str(item.get("extension", "")).lower() == "xlsx" and (
                series_requirement_keyword(item["name"]) == target_requirement
                if target_requirement
                else target_keyword in item["name"]
            ):
                mode = ""
                version = ""
                path_parts = item["path"].split("/")

                for index, sp in enumerate(path_parts):
                    is_target_file = "xlsx" in sp.lower() and (
                        series_requirement_keyword(sp) == target_requirement
                        if target_requirement
                        else target_keyword in sp
                    )
                    if is_target_file:
                        if target_keyword == "外设功能分析表":
                            parent = path_parts[index - 1] if index > 0 else ""
                            mode = parent if parent and target_keyword not in parent else mode
                        elif uses_analysis_version:
                            mode = target_mode
                        if uses_analysis_version:
                            found = re.search(r"v(\d+(?:\.\d+)*)", sp.lower())
                        else:
                            mode = re.split(r"[_-]", sp)[0] if mode != "common" else mode
                            found = re.search(r"v(\d+\.\d+)\.xlsx", sp.lower())
                        if found:
                            version = found.group(1)
                        elif uses_analysis_version:
                            version = "0"
                        else:
                            print(f"{item['path']}的版本号不正确")
                    elif sp.lower() == "common":
                        mode = "common"

                if mode.lower() != target_mode.lower():
                    continue

                url_respone = _api.post(f"/v1.0/storage/spaces/{SPACE_ID}/dentries/{item['id']}/downloadInfos/query",
                                        headers={"x-acs-dingtalk-access-token":accessToken},
                                        params={"unionId": UNION_ID})
                url_respone.raise_for_status()
                headers = url_respone.json()["headerSignatureInfo"]["headers"]
                resourceUrl = url_respone.json()["headerSignatureInfo"]["resourceUrls"][0]

                form_key = f"{series}_{mode}"
                subdirectory = ""
                if collect_all_change_forms:
                    mode_indexes = [
                        index
                        for index, part in enumerate(path_parts[:-1])
                        if part.casefold() == target_mode.casefold()
                    ]
                    if mode_indexes and mode_indexes[-1] + 1 < len(path_parts) - 1:
                        subdirectory = path_parts[mode_indexes[-1] + 1]
                        form_key += f"_{subdirectory}"
                if form_key not in change_application_form or _version_key(version) > _version_key(change_application_form[form_key]["version"]):
                    change_application_form[form_key] = {
                        "mode": mode,
                        "headers": headers,
                        "resourceUrl": resourceUrl,
                        "version": version,
                        "name": item["name"],
                        "subdirectory": subdirectory,
                    }
    if len(change_application_form) > 0 and not collect_all_change_forms:
        return

def download_file(path):
    for series in change_application_form:
        d_path = os.path.join(
            path,
            change_application_form[series].get("subdirectory", ""),
            change_application_form[series]["name"],
        )
        os.makedirs(os.path.dirname(d_path), exist_ok=True)

        with requests.get(change_application_form[series]['resourceUrl'], headers=change_application_form[series]['headers'], stream=True, timeout=30) as response:
            response.raise_for_status()  # 如果响应状态码不是 200，抛出异常

            # 将返回的二进制内容写入本地文件
            with open(d_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)


# TODO: common模块的特殊处理
def get_change_application_form(series, mode, path):
    change_application_form.clear()
    accessToken = get_access_token()
    for id in folder_ids:
        get_all_files(accessToken, id, series, mode)
    download_file(path)

def get_peripheral_feature_analysis_form(series, mode, path, target_keyword="外设功能分析表"):
    change_application_form.clear()
    accessToken = get_access_token()
    for id in folder_ids:
        get_all_files(accessToken, id, series, mode, target_keyword=target_keyword)
    download_file(path)
