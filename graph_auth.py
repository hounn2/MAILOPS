"""
Microsoft Graph API认证模块
支持设备代码流和设备授权流
"""

import os
import json
import logging
import msal
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


class GraphAuth:
    """Microsoft Graph API认证管理器"""

    # Microsoft Graph API端点
    AUTHORITY = "https://login.microsoftonline.com/common"
    SCOPE = ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "User.Read"]

    def __init__(self, client_id: str, token_cache_path: str = "token_cache.json"):
        """
        初始化认证管理器

        Args:
            client_id: Azure AD应用客户端ID
            token_cache_path: Token缓存文件路径
        """
        self.client_id = client_id
        self.token_cache_path = token_cache_path
        self.access_token = None
        self.app = None
        self._init_app()

    def _init_app(self):
        """初始化MSAL应用"""
        try:
            # 加载或创建token缓存
            cache = msal.SerializableTokenCache()

            if os.path.exists(self.token_cache_path):
                with open(self.token_cache_path, "r") as f:
                    cache.deserialize(f.read())

            self.app = msal.PublicClientApplication(
                client_id=self.client_id, authority=self.AUTHORITY, token_cache=cache
            )

            # 保存缓存的回调
            self.cache = cache

            logger.info("MSAL应用初始化成功")

        except Exception as e:
            logger.error(f"初始化MSAL应用失败: {e}")
            raise

    def _save_cache(self):
        """保存token缓存到文件"""
        try:
            if self.cache.has_state_changed:
                with open(self.token_cache_path, "w") as f:
                    f.write(self.cache.serialize())
        except Exception as e:
            logger.error(f"保存token缓存失败: {e}")

    def authenticate(self) -> bool:
        """
        执行认证流程

        Returns:
            是否认证成功
        """
        try:
            # 首先尝试从缓存获取账户
            accounts = self.app.get_accounts()

            if accounts:
                # 尝试静默获取token
                result = self.app.acquire_token_silent(self.SCOPE, account=accounts[0])
                if result:
                    self.access_token = result["access_token"]
                    logger.info("从缓存成功获取token")
                    self._save_cache()
                    return True

            # 如果没有缓存或token过期，使用设备代码流
            return self._device_code_flow()

        except Exception as e:
            logger.error(f"认证失败: {e}")
            return False

    def _device_code_flow(self) -> bool:
        """
        使用设备代码流进行认证

        Returns:
            是否认证成功
        """
        try:
            # 获取设备代码
            flow = self.app.initiate_device_flow(scopes=self.SCOPE)

            if "user_code" not in flow:
                logger.error("无法创建设备代码流")
                return False

            # 显示给用户
            print("\n" + "=" * 60)
            print("需要进行Microsoft账户认证")
            print("=" * 60)
            print(f"1. 在浏览器中打开: {flow['verification_uri']}")
            print(f"2. 输入代码: {flow['user_code']}")
            print("=" * 60 + "\n")

            # 等待用户完成认证
            result = self.app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("设备代码流认证成功")
                self._save_cache()
                return True
            else:
                error = result.get("error_description", "未知错误")
                logger.error(f"认证失败: {error}")
                return False

        except Exception as e:
            logger.error(f"设备代码流认证失败: {e}")
            return False

    def get_token(self) -> Optional[str]:
        """
        获取access token

        Returns:
            access token或None
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        return self.access_token

    def get_headers(self) -> Dict[str, str]:
        """
        获取HTTP请求头

        Returns:
            包含Authorization的请求头
        """
        token = self.get_token()
        if not token:
            raise Exception("无法获取有效的access token")

        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    def logout(self):
        """注销并清除缓存"""
        try:
            if os.path.exists(self.token_cache_path):
                os.remove(self.token_cache_path)
                logger.info("已清除token缓存")

            self.access_token = None
            self._init_app()  # 重新初始化

        except Exception as e:
            logger.error(f"注销失败: {e}")


class GraphAuthInteractive:
    """交互式认证（适用于有GUI的环境）"""

    AUTHORITY = "https://login.microsoftonline.com/common"
    SCOPE = ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "User.Read"]

    def __init__(self, client_id: str, redirect_uri: str = "http://localhost:8080"):
        """
        初始化交互式认证

        Args:
            client_id: Azure AD应用客户端ID
            redirect_uri: 重定向URI
        """
        self.client_id = client_id
        self.redirect_uri = redirect_uri
        self.access_token = None

        self.app = msal.PublicClientApplication(
            client_id=client_id, authority=self.AUTHORITY
        )

    def authenticate(self) -> bool:
        """
        使用交互式登录

        Returns:
            是否认证成功
        """
        try:
            # 尝试获取已有账户
            accounts = self.app.get_accounts()
            if accounts:
                result = self.app.acquire_token_silent(self.SCOPE, account=accounts[0])
                if result:
                    self.access_token = result["access_token"]
                    return True

            # 交互式登录
            result = self.app.acquire_token_interactive(
                scopes=self.SCOPE, redirect_uri=self.redirect_uri
            )

            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("交互式认证成功")
                return True
            else:
                error = result.get("error_description", "未知错误")
                logger.error(f"认证失败: {error}")
                return False

        except Exception as e:
            logger.error(f"交互式认证失败: {e}")
            return False

    def get_token(self) -> Optional[str]:
        """获取access token"""
        if not self.access_token:
            if not self.authenticate():
                return None
        return self.access_token

    def get_headers(self) -> Dict[str, str]:
        """获取HTTP请求头"""
        token = self.get_token()
        if not token:
            raise Exception("无法获取有效的access token")

        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
