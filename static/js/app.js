// 全局变量
let currentRules = [];
let currentPage = 'dashboard';
let ruleModal = null;
let logDetailModal = null;

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    ruleModal = new bootstrap.Modal(document.getElementById('ruleModal'));
    logDetailModal = new bootstrap.Modal(document.getElementById('logDetailModal'));
    
    // 绑定菜单点击事件
    document.querySelectorAll('.list-group-item').forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            const page = this.dataset.page;
            showPage(page);
        });
    });
    
    // 初始化页面
    loadDashboard();
    setInterval(loadDashboard, 30000); // 每30秒刷新一次
});

// 页面切换
function showPage(page) {
    // 隐藏所有页面
    document.querySelectorAll('.page-content').forEach(el => {
        el.style.display = 'none';
    });
    
    // 显示目标页面
    document.getElementById(page + '-page').style.display = 'block';
    
    // 更新菜单激活状态
    document.querySelectorAll('.list-group-item').forEach(el => {
        el.classList.remove('active');
    });
    document.querySelector(`[data-page="${page}"]`).classList.add('active');
    
    currentPage = page;
    
    // 加载页面数据
    switch(page) {
        case 'dashboard':
            loadDashboard();
            break;
        case 'rules':
            loadRules();
            break;
        case 'logs':
            loadLogs();
            break;
        case 'settings':
            loadSettings();
            break;
    }
}

// ==================== 仪表盘 ====================

function loadDashboard() {
    // 加载统计数据
    fetch('/api/statistics')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const stats = data.statistics;
                document.getElementById('totalExecutions').textContent = stats.total_executions;
                document.getElementById('totalEmails').textContent = stats.total_emails;
                document.getElementById('totalMatched').textContent = stats.total_matched;
                document.getElementById('totalActions').textContent = stats.total_actions;
                document.getElementById('todayEmails').textContent = stats.today_emails;
                document.getElementById('todayMatched').textContent = stats.today_matched;
                document.getElementById('todayActions').textContent = stats.today_actions;
            }
        });
    
    // 加载最近记录
    loadRecentLogs();
}

function loadRecentLogs() {
    fetch('/api/logs?limit=5')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const tbody = document.getElementById('recentLogs');
                if (data.logs.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">暂无记录</td></tr>';
                    return;
                }
                
                tbody.innerHTML = data.logs.map(log => `
                    <tr>
                        <td>${formatDateTime(log.execution_time)}</td>
                        <td>${log.total_emails}</td>
                        <td>${log.matched_emails}</td>
                        <td>${log.actions_executed}</td>
                        <td>${log.duration.toFixed(1)}s</td>
                        <td>
                            <span class="badge bg-${log.status === 'success' ? 'success' : 'danger'}">
                                ${log.status === 'success' ? '成功' : '失败'}
                            </span>
                        </td>
                    </tr>
                `).join('');
            }
        });
}

// ==================== 规则管理 ====================

function loadRules() {
    fetch('/api/rules')
        .then(response => response.json())
        .then(rules => {
            currentRules = rules;
            const container = document.getElementById('rulesList');
            
            if (rules.length === 0) {
                container.innerHTML = `
                    <div class="alert alert-info">
                        <i class="bi bi-info-circle"></i> 暂无规则，点击"新建规则"创建第一个规则。
                    </div>
                `;
                return;
            }
            
            container.innerHTML = rules.map((rule, index) => `
                <div class="rule-item ${rule.enabled ? '' : 'disabled'}">
                    <div class="d-flex justify-content-between align-items-start">
                        <div>
                            <h5 class="mb-1">
                                ${index + 1}. ${rule.name}
                                ${rule.enabled ? 
                                    '<span class="badge bg-success">已启用</span>' : 
                                    '<span class="badge bg-secondary">已禁用</span>'}
                            </h5>
                            <small class="text-muted">ID: ${rule.id}</small>
                        </div>
                        <div class="btn-group">
                            <button class="btn btn-sm btn-outline-primary" onclick="editRule('${rule.id}')">
                                <i class="bi bi-pencil"></i> 编辑
                            </button>
                            <button class="btn btn-sm btn-outline-danger" onclick="deleteRule('${rule.id}')">
                                <i class="bi bi-trash"></i> 删除
                            </button>
                        </div>
                    </div>
                    <div class="mt-3">
                        <h6>条件 (${rule.conditions.match_all ? '全部满足' : '任一满足'}):</h6>
                        <ul class="list-unstyled">
                            ${rule.conditions.items.map(item => `
                                <li><i class="bi bi-check-circle-fill text-primary"></i> 
                                    ${getFieldName(item.field)} ${getOperatorName(item.operator)} 
                                    ${Array.isArray(item.value) ? item.value.join(' / ') : item.value}
                                </li>
                            `).join('')}
                        </ul>
                        <h6>动作:</h6>
                        <ul class="list-unstyled">
                            ${rule.actions.map(action => `
                                <li><i class="bi bi-arrow-right-circle-fill text-success"></i> 
                                    ${getActionName(action)}
                                </li>
                            `).join('')}
                        </ul>
                    </div>
                </div>
            `).join('');
        });
}

function getFieldName(field) {
    const names = {
        'subject': '主题',
        'body': '正文',
        'sender': '发件人',
        'sender_domain': '发件人域名',
        'has_attachments': '有附件',
        'received_time': '接收时间'
    };
    return names[field] || field;
}

function getOperatorName(operator) {
    const names = {
        'equals': '等于',
        'not_equals': '不等于',
        'contains': '包含',
        'not_contains': '不包含',
        'starts_with': '开头是',
        'ends_with': '结尾是',
        'in': '在列表中',
        'not_in': '不在列表中'
    };
    return names[operator] || operator;
}

function getActionName(action) {
    const names = {
        'reply': `回复邮件${action.template ? ' (使用模板: ' + action.template + ')' : ''}`,
        'forward': `转发给 ${action.to ? action.to.join(', ') : ''}`,
        'move': `移动到 "${action.target}"`,
        'mark_as_read': '标记为已读',
        'ai_reply': '🤖 AI智能回复' + (action.use_knowledge_base !== false ? ' (使用知识库)' : '')
    };
    return names[action.type] || action.type;
}

// ==================== 规则编辑 ====================

function addNewRule() {
    document.getElementById('ruleForm').reset();
    document.getElementById('ruleId').value = '';
    document.getElementById('conditionsList').innerHTML = '';
    document.getElementById('actionsList').innerHTML = '';
    
    // 添加默认条件和动作
    addCondition();
    addAction();
    
    ruleModal.show();
}

function editRule(ruleId) {
    const rule = currentRules.find(r => r.id === ruleId);
    if (!rule) return;
    
    document.getElementById('ruleId').value = rule.id;
    document.getElementById('ruleName').value = rule.name;
    document.getElementById('ruleEnabled').checked = rule.enabled;
    document.getElementById('matchAll').value = rule.conditions.match_all.toString();
    
    // 填充条件
    document.getElementById('conditionsList').innerHTML = '';
    rule.conditions.items.forEach(item => addCondition(item));
    
    // 填充动作
    document.getElementById('actionsList').innerHTML = '';
    rule.actions.forEach(action => addAction(action));
    
    ruleModal.show();
}

function addCondition(data = null) {
    const div = document.createElement('div');
    div.className = 'condition-row';
    div.innerHTML = `
        <div class="row">
            <div class="col-md-3">
                <select class="form-select condition-field">
                    <option value="subject" ${data?.field === 'subject' ? 'selected' : ''}>主题</option>
                    <option value="body" ${data?.field === 'body' ? 'selected' : ''}>正文</option>
                    <option value="sender" ${data?.field === 'sender' ? 'selected' : ''}>发件人</option>
                    <option value="sender_domain" ${data?.field === 'sender_domain' ? 'selected' : ''}>发件人域名</option>
                    <option value="has_attachments" ${data?.field === 'has_attachments' ? 'selected' : ''}>有附件</option>
                </select>
            </div>
            <div class="col-md-3">
                <select class="form-select condition-operator">
                    <option value="equals" ${data?.operator === 'equals' ? 'selected' : ''}>等于</option>
                    <option value="not_equals" ${data?.operator === 'not_equals' ? 'selected' : ''}>不等于</option>
                    <option value="contains" ${data?.operator === 'contains' ? 'selected' : ''}>包含</option>
                    <option value="not_contains" ${data?.operator === 'not_contains' ? 'selected' : ''}>不包含</option>
                    <option value="starts_with" ${data?.operator === 'starts_with' ? 'selected' : ''}>开头是</option>
                    <option value="ends_with" ${data?.operator === 'ends_with' ? 'selected' : ''}>结尾是</option>
                    <option value="in" ${data?.operator === 'in' ? 'selected' : ''}>在列表中</option>
                    <option value="not_in" ${data?.operator === 'not_in' ? 'selected' : ''}>不在列表中</option>
                </select>
            </div>
            <div class="col-md-5">
                <input type="text" class="form-control condition-value" 
                       placeholder="多个值用逗号分隔" 
                       value="${data ? (Array.isArray(data.value) ? data.value.join(', ') : data.value) : ''}">
            </div>
            <div class="col-md-1">
                <button type="button" class="btn btn-sm btn-outline-danger" onclick="this.closest('.condition-row').remove()">
                    <i class="bi bi-trash"></i>
                </button>
            </div>
        </div>
    `;
    document.getElementById('conditionsList').appendChild(div);
}

function addAction(data = null) {
    const div = document.createElement('div');
    div.className = 'action-row';
    div.innerHTML = `
        <div class="row">
            <div class="col-md-3">
                <select class="form-select action-type" onchange="updateActionFields(this)">
                    <option value="reply" ${data?.type === 'reply' ? 'selected' : ''}>回复邮件</option>
                    <option value="forward" ${data?.type === 'forward' ? 'selected' : ''}>转发邮件</option>
                    <option value="move" ${data?.type === 'move' ? 'selected' : ''}>移动邮件</option>
                    <option value="mark_as_read" ${data?.type === 'mark_as_read' ? 'selected' : ''}>标记为已读</option>
                    <option value="ai_reply" ${data?.type === 'ai_reply' ? 'selected' : ''}>🤖 AI智能回复</option>
                </select>
            </div>
            <div class="col-md-8 action-fields">
                ${getActionFields(data)}
            </div>
            <div class="col-md-1">
                <button type="button" class="btn btn-sm btn-outline-danger" onclick="this.closest('.action-row').remove()">
                    <i class="bi bi-trash"></i>
                </button>
            </div>
        </div>
    `;
    document.getElementById('actionsList').appendChild(div);
}

function getActionFields(data) {
    if (!data) {
        return `
            <input type="text" class="form-control action-template" placeholder="模板名称">
            <div class="form-check mt-2">
                <input class="form-check-input action-include-original" type="checkbox">
                <label class="form-check-label">包含原始邮件</label>
            </div>
        `;
    }
    
    switch(data.type) {
        case 'reply':
            return `
                <input type="text" class="form-control action-template" placeholder="模板名称" value="${data.template || ''}">
                <div class="form-check mt-2">
                    <input class="form-check-input action-include-original" type="checkbox" ${data.include_original ? 'checked' : ''}>
                    <label class="form-check-label">包含原始邮件</label>
                </div>
            `;
        case 'forward':
            return `
                <input type="text" class="form-control action-to" placeholder="收件人邮箱（多个用逗号分隔）" value="${data.to ? data.to.join(', ') : ''}">
                <input type="text" class="form-control action-prefix mt-2" placeholder="主题前缀" value="${data.subject_prefix || ''}">
            `;
        case 'move':
            return `
                <input type="text" class="form-control action-target" placeholder="目标文件夹名称" value="${data.target || ''}">
            `;
        case 'ai_reply':
            return `
                <div class="form-check mb-2">
                    <input class="form-check-input action-use-kb" type="checkbox" ${data.use_knowledge_base !== false ? 'checked' : ''}>
                    <label class="form-check-label">使用知识库</label>
                </div>
                <div class="form-check mb-2">
                    <input class="form-check-input action-include-original" type="checkbox" ${data.include_original ? 'checked' : ''}>
                    <label class="form-check-label">包含原始邮件</label>
                </div>
                <input type="text" class="form-control action-subject mb-2" placeholder="回复主题模板" value="${data.subject || 'RE: {original_subject}'}">
                <input type="number" class="form-control action-temperature" placeholder="温度参数(0-1)" value="${data.temperature || 0.7}" min="0" max="1" step="0.1">
            `;
        default:
            return '';
    }
}

function updateActionFields(select) {
    const row = select.closest('.action-row');
    const fieldsDiv = row.querySelector('.action-fields');
    const type = select.value;
    
    switch(type) {
        case 'reply':
            fieldsDiv.innerHTML = `
                <input type="text" class="form-control action-template" placeholder="模板名称">
                <div class="form-check mt-2">
                    <input class="form-check-input action-include-original" type="checkbox">
                    <label class="form-check-label">包含原始邮件</label>
                </div>
            `;
            break;
        case 'forward':
            fieldsDiv.innerHTML = `
                <input type="text" class="form-control action-to" placeholder="收件人邮箱（多个用逗号分隔）">
                <input type="text" class="form-control action-prefix mt-2" placeholder="主题前缀">
            `;
            break;
        case 'move':
            fieldsDiv.innerHTML = `
                <input type="text" class="form-control action-target" placeholder="目标文件夹名称">
            `;
            break;
        case 'mark_as_read':
            fieldsDiv.innerHTML = '<span class="text-muted">无需额外配置</span>';
            break;
        case 'ai_reply':
            fieldsDiv.innerHTML = `
                <div class="form-check mb-2">
                    <input class="form-check-input action-use-kb" type="checkbox" checked>
                    <label class="form-check-label">使用知识库</label>
                </div>
                <div class="form-check mb-2">
                    <input class="form-check-input action-include-original" type="checkbox">
                    <label class="form-check-label">包含原始邮件</label>
                </div>
                <input type="text" class="form-control action-subject mb-2" placeholder="回复主题模板" value="RE: {original_subject}">
                <input type="number" class="form-control action-temperature" placeholder="温度参数(0-1)" value="0.7" min="0" max="1" step="0.1">
            `;
            break;
    }
}

function saveRule() {
    const ruleId = document.getElementById('ruleId').value;
    const rule = {
        id: ruleId || undefined,
        name: document.getElementById('ruleName').value,
        enabled: document.getElementById('ruleEnabled').checked,
        conditions: {
            match_all: document.getElementById('matchAll').value === 'true',
            items: []
        },
        actions: []
    };
    
    // 收集条件
    document.querySelectorAll('.condition-row').forEach(row => {
        const field = row.querySelector('.condition-field').value;
        const operator = row.querySelector('.condition-operator').value;
        let value = row.querySelector('.condition-value').value;
        
        // 如果是列表操作符，转换为数组
        if (['in', 'not_in', 'contains', 'not_contains'].includes(operator)) {
            value = value.split(/[,，]/).map(v => v.trim()).filter(v => v);
        }
        
        rule.conditions.items.push({ field, operator, value });
    });
    
    // 收集动作
    document.querySelectorAll('.action-row').forEach(row => {
        const type = row.querySelector('.action-type').value;
        const action = { type };
        
        switch(type) {
            case 'reply':
                const template = row.querySelector('.action-template')?.value;
                const includeOriginal = row.querySelector('.action-include-original')?.checked;
                if (template) action.template = template;
                if (includeOriginal) action.include_original = true;
                break;
            case 'forward':
                const to = row.querySelector('.action-to')?.value;
                const prefix = row.querySelector('.action-prefix')?.value;
                if (to) action.to = to.split(/[,，]/).map(v => v.trim()).filter(v => v);
                if (prefix) action.subject_prefix = prefix;
                break;
            case 'move':
                const target = row.querySelector('.action-target')?.value;
                if (target) action.target = target;
                break;
            case 'ai_reply':
                const useKB = row.querySelector('.action-use-kb')?.checked;
                const includeOrig = row.querySelector('.action-include-original')?.checked;
                const subjectTpl = row.querySelector('.action-subject')?.value;
                const temperature = row.querySelector('.action-temperature')?.value;
                action.use_knowledge_base = useKB !== false;
                if (includeOrig) action.include_original = true;
                if (subjectTpl) action.subject = subjectTpl;
                if (temperature) action.temperature = parseFloat(temperature);
                break;
        }

        rule.actions.push(action);
    });
    
    const url = ruleId ? `/api/rules/${ruleId}` : '/api/rules';
    const method = ruleId ? 'PUT' : 'POST';
    
    fetch(url, {
        method: method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(rule)
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            ruleModal.hide();
            loadRules();
            showAlert('规则保存成功', 'success');
        } else {
            showAlert('保存失败: ' + data.message, 'danger');
        }
    })
    .catch(error => {
        showAlert('保存失败: ' + error, 'danger');
    });
}

function deleteRule(ruleId) {
    if (!confirm('确定要删除这条规则吗？')) return;
    
    fetch(`/api/rules/${ruleId}`, {
        method: 'DELETE'
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            loadRules();
            showAlert('规则已删除', 'success');
        } else {
            showAlert('删除失败', 'danger');
        }
    });
}

// ==================== 执行记录 ====================

function loadLogs() {
    fetch('/api/logs?limit=50')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const tbody = document.getElementById('allLogs');
                if (data.logs.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted">暂无记录</td></tr>';
                    return;
                }
                
                tbody.innerHTML = data.logs.map(log => `
                    <tr>
                        <td>${formatDateTime(log.execution_time)}</td>
                        <td>${log.total_emails}</td>
                        <td>${log.matched_emails}</td>
                        <td>${log.actions_executed}</td>
                        <td>${log.errors}</td>
                        <td>${log.duration.toFixed(1)}s</td>
                        <td>
                            <span class="badge bg-${log.status === 'success' ? 'success' : 'danger'}">
                                ${log.status === 'success' ? '成功' : '失败'}
                            </span>
                        </td>
                        <td>
                            <button class="btn btn-sm btn-outline-info" onclick="showLogDetail(${log.id})">
                                <i class="bi bi-eye"></i> 详情
                            </button>
                        </td>
                    </tr>
                `).join('');
            }
        });
}

function showLogDetail(logId) {
    console.log('Loading details for log ID:', logId);
    fetch(`/api/logs/${logId}/details`)
        .then(response => response.json())
        .then(data => {
            console.log('API Response:', data);
            const content = document.getElementById('logDetailContent');
            if (!data.success) {
                content.innerHTML = `<p class="text-danger">加载失败: ${data.message}</p>`;
                logDetailModal.show();
                return;
            }
            
            if (!data.details || data.details.length === 0) {
                content.innerHTML = '<p class="text-muted">暂无详细记录</p>';
            } else {
                content.innerHTML = `
                    <table class="table table-sm">
                        <thead>
                            <tr>
                                <th>邮件主题</th>
                                <th>发件人</th>
                                <th>匹配规则</th>
                                <th>执行动作</th>
                                <th>状态</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${data.details.map(d => {
                                // 安全地解析JSON
                                let matchedRules = [];
                                let actionsTaken = [];
                                try {
                                    matchedRules = JSON.parse(d.matched_rules || '[]');
                                } catch(e) { matchedRules = []; }
                                try {
                                    actionsTaken = JSON.parse(d.actions_taken || '[]');
                                } catch(e) { actionsTaken = []; }
                                
                                return `
                                    <tr>
                                        <td>${d.email_subject || '-'}</td>
                                        <td>${d.sender || '-'}</td>
                                        <td>${matchedRules.join(', ') || '-'}</td>
                                        <td>${actionsTaken.join(', ') || '-'}</td>
                                        <td><span class="badge bg-secondary">${d.status || 'unknown'}</span></td>
                                    </tr>
                                `;
                            }).join('')}
                        </tbody>
                    </table>
                `;
            }
            logDetailModal.show();
        })
        .catch(error => {
            console.error('Error loading details:', error);
            document.getElementById('logDetailContent').innerHTML = `<p class="text-danger">加载失败: ${error}</p>`;
            logDetailModal.show();
        });
}

// ==================== 系统设置 ====================

function loadSettings() {
    fetch('/api/config')
        .then(response => response.json())
        .then(config => {
            const settings = config.settings || {};
            document.getElementById('checkInterval').value = settings.check_interval || 60;
            document.getElementById('processUnreadOnly').checked = settings.process_unread_only !== false;
            document.getElementById('maxEmailsPerBatch').value = settings.max_emails_per_batch || 50;
            document.getElementById('markAsRead').checked = settings.mark_as_read_after_process === true;
        });
    
    // 加载自动执行状态
    loadAutoExecutionStatus();
}

function loadAutoExecutionStatus() {
    fetch('/api/auto_execution')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const checkbox = document.getElementById('autoExecutionEnabled');
                const statusDiv = document.getElementById('autoExecutionStatus');
                const statusText = document.getElementById('autoExecutionText');
                
                checkbox.checked = data.enabled;
                
                if (data.enabled && data.running) {
                    statusDiv.style.display = 'block';
                    statusDiv.className = 'alert alert-success';
                    statusText.textContent = `自动执行运行中（每${data.interval}秒检查一次）`;
                } else if (data.enabled && !data.running) {
                    statusDiv.style.display = 'block';
                    statusDiv.className = 'alert alert-warning';
                    statusText.textContent = '自动执行已启用但未运行，请重启服务';
                } else {
                    statusDiv.style.display = 'none';
                }
            }
        });
}

function toggleAutoExecution() {
    const checkbox = document.getElementById('autoExecutionEnabled');
    const enabled = checkbox.checked;
    
    fetch('/api/auto_execution', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ enabled: enabled })
    })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert(data.message, 'success');
                loadAutoExecutionStatus();
            } else {
                showAlert(data.message, 'danger');
                checkbox.checked = !enabled; // 恢复状态
            }
        })
        .catch(error => {
            showAlert('设置失败: ' + error, 'danger');
            checkbox.checked = !enabled; // 恢复状态
        });
}

function saveSettings() {
    fetch('/api/config')
        .then(response => response.json())
        .then(config => {
            config.settings = config.settings || {};
            config.settings.check_interval = parseInt(document.getElementById('checkInterval').value);
            config.settings.process_unread_only = document.getElementById('processUnreadOnly').checked;
            config.settings.max_emails_per_batch = parseInt(document.getElementById('maxEmailsPerBatch').value);
            config.settings.mark_as_read_after_process = document.getElementById('markAsRead').checked;
            config.settings.auto_execution = document.getElementById('autoExecutionEnabled').checked;
            
            return fetch('/api/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(config)
            });
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert('设置已保存', 'success');
            } else {
                showAlert('保存失败', 'danger');
            }
        });
}

// ==================== 执行控制 ====================

function executeNow() {
    const btn = document.getElementById('executeBtn');
    const dryRun = document.getElementById('dryRunCheck').checked;
    
    btn.disabled = true;
    btn.innerHTML = '<i class="bi bi-hourglass-split"></i> 执行中...';
    
    document.getElementById('statusText').innerHTML = 
        '<i class="bi bi-circle-fill text-warning status-processing"></i> 执行中';
    
    fetch('/api/execute', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ dry_run: dryRun })
    })
    .then(response => response.json())
    .then(data => {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-play-fill"></i> 立即执行';
        document.getElementById('statusText').innerHTML = 
            '<i class="bi bi-circle-fill text-success"></i> 就绪';
        
        if (data.success) {
            showAlert(data.message, 'success');
            // 刷新当前页面数据
            if (currentPage === 'dashboard') loadDashboard();
            if (currentPage === 'logs') loadLogs();
        } else {
            showAlert(data.message, 'danger');
        }
    })
    .catch(error => {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-play-fill"></i> 立即执行';
        document.getElementById('statusText').innerHTML = 
            '<i class="bi bi-circle-fill text-success"></i> 就绪';
        showAlert('执行失败: ' + error, 'danger');
    });
}

// ==================== 工具函数 ====================

function formatDateTime(isoString) {
    if (!isoString) return '-';
    const date = new Date(isoString);
    return date.toLocaleString('zh-CN');
}

function showAlert(message, type) {
    const div = document.createElement('div');
    div.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
    div.style.cssText = 'top: 20px; right: 20px; z-index: 9999; min-width: 300px;';
    div.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    document.body.appendChild(div);
    
    setTimeout(() => {
        div.remove();
    }, 5000);
}

// ==================== AI功能 ====================

function loadAIPage() {
    // 加载AI状态
    fetch('/api/ai/status')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // 更新AI引擎状态
                document.getElementById('aiEnabled').checked = data.ai.enabled;
                document.getElementById('aiBaseUrl').value = data.ai.config.base_url || 'http://localhost:1234';
                document.getElementById('aiModel').value = data.ai.config.model || '';
                document.getElementById('aiTimeout').value = data.ai.config.timeout || 60;
                document.getElementById('aiSystemPrompt').value = data.ai.config.system_prompt || '';
                
                const aiBadge = document.getElementById('aiStatusBadge');
                if (data.ai.enabled) {
                    if (data.ai.ready) {
                        aiBadge.className = 'badge bg-success';
                        aiBadge.textContent = '运行中';
                    } else {
                        aiBadge.className = 'badge bg-warning';
                        aiBadge.textContent = data.ai.error || '未连接';
                    }
                } else {
                    aiBadge.className = 'badge bg-secondary';
                    aiBadge.textContent = '未启用';
                }
                
                toggleAI();
                
                // 更新知识库状态
                document.getElementById('kbEnabled').checked = data.knowledge_base.enabled;
                document.getElementById('kbPath').value = data.knowledge_base.config.path || 'knowledge_base';
                document.getElementById('kbTopK').value = data.knowledge_base.config.search_top_k || 3;
                
                const kbBadge = document.getElementById('kbStatusBadge');
                if (data.knowledge_base.enabled) {
                    if (data.knowledge_base.ready) {
                        kbBadge.className = 'badge bg-success';
                        kbBadge.textContent = `${data.knowledge_base.stats.total_documents} 个文档`;
                    } else {
                        kbBadge.className = 'badge bg-warning';
                        kbBadge.textContent = '未就绪';
                    }
                } else {
                    kbBadge.className = 'badge bg-secondary';
                    kbBadge.textContent = '未启用';
                }
                
                toggleKB();
                
                // 显示测试区域
                if (data.ai.enabled && data.ai.ready) {
                    document.getElementById('aiTestSection').style.display = 'block';
                } else {
                    document.getElementById('aiTestSection').style.display = 'none';
                }
            }
        });
}

function toggleAI() {
    const enabled = document.getElementById('aiEnabled').checked;
    document.getElementById('aiConfigSection').style.display = enabled ? 'block' : 'none';
}

function toggleKB() {
    const enabled = document.getElementById('kbEnabled').checked;
    document.getElementById('kbConfigSection').style.display = enabled ? 'block' : 'none';
}

function saveAIConfig() {
    fetch('/api/config')
        .then(response => response.json())
        .then(config => {
            config.lmstudio = {
                enabled: document.getElementById('aiEnabled').checked,
                base_url: document.getElementById('aiBaseUrl').value,
                model: document.getElementById('aiModel').value,
                timeout: parseInt(document.getElementById('aiTimeout').value),
                system_prompt: document.getElementById('aiSystemPrompt').value
            };
            
            return fetch('/api/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(config)
            });
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert('AI配置已保存', 'success');
                loadAIPage();
            } else {
                showAlert('保存失败: ' + data.message, 'danger');
            }
        });
}

function saveKBConfig() {
    fetch('/api/config')
        .then(response => response.json())
        .then(config => {
            config.knowledge_base = {
                enabled: document.getElementById('kbEnabled').checked,
                path: document.getElementById('kbPath').value,
                search_top_k: parseInt(document.getElementById('kbTopK').value)
            };
            
            return fetch('/api/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(config)
            });
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert('知识库配置已保存', 'success');
                loadAIPage();
            } else {
                showAlert('保存失败: ' + data.message, 'danger');
            }
        });
}

function testAIConnection() {
    const btn = event.target;
    btn.disabled = true;
    btn.innerHTML = '<i class="bi bi-hourglass-split"></i> 测试中...';
    
    // 获取当前输入框中的配置
    const baseUrl = document.getElementById('aiBaseUrl').value || 'http://localhost:1234';
    
    // 通过后端代理测试连接，避免跨域问题
    fetch('/api/ai/status?base_url=' + encodeURIComponent(baseUrl))
        .then(response => response.json())
        .then(data => {
            btn.disabled = false;
            btn.innerHTML = '测试连接';
            
            if (data.success && data.ai.ready) {
                showAlert(`LMStudio连接成功！可用模型: ${data.ai.model_count}个`, 'success');
            } else {
                showAlert('连接失败: ' + (data.ai.error || '请检查LMStudio是否运行'), 'danger');
            }
        })
        .catch(error => {
            btn.disabled = false;
            btn.innerHTML = '测试连接';
            showAlert('测试失败: ' + error, 'danger');
        });
}

function testAIGenerate() {
    const content = document.getElementById('testEmailContent').value;
    if (!content.trim()) {
        showAlert('请输入测试内容', 'warning');
        return;
    }
    
    const btn = event.target;
    btn.disabled = true;
    btn.innerHTML = '<i class="bi bi-hourglass-split"></i> 生成中...';
    
    fetch('/api/ai/test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            email_content: content,
            use_knowledge_base: document.getElementById('testUseKB').checked
        })
    })
        .then(response => response.json())
        .then(data => {
            btn.disabled = false;
            btn.innerHTML = '<i class="bi bi-play-fill"></i> 生成回复';
            
            const resultDiv = document.getElementById('aiTestResult');
            const replyDiv = document.getElementById('aiTestReply');
            
            if (data.success) {
                replyDiv.textContent = data.reply;
                resultDiv.style.display = 'block';
            } else {
                showAlert('生成失败: ' + data.message, 'danger');
            }
        })
        .catch(error => {
            btn.disabled = false;
            btn.innerHTML = '<i class="bi bi-play-fill"></i> 生成回复';
            showAlert('请求失败: ' + error, 'danger');
        });
}

// ==================== 知识库文件管理 ====================

function loadKBFiles() {
    fetch('/api/kb/files')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const tbody = document.querySelector('#kbFilesTable tbody');
                if (data.files.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted">暂无文件</td></tr>';
                    return;
                }
                
                tbody.innerHTML = data.files.map(file => `
                    <tr>
                        <td>${file.name}</td>
                        <td>${formatFileSize(file.size)}</td>
                        <td>${formatDateTime(file.modified)}</td>
                        <td>
                            <button class="btn btn-sm btn-outline-danger" onclick="deleteKBFile('${file.name}')">
                                <i class="bi bi-trash"></i> 删除
                            </button>
                        </td>
                    </tr>
                `).join('');
            }
        });
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function uploadKBFile(input) {
    const file = input.files[0];
    if (!file) return;
    
    const formData = new FormData();
    formData.append('file', file);
    
    fetch('/api/kb/upload', {
        method: 'POST',
        body: formData
    })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert(data.message, 'success');
                loadKBFiles();
            } else {
                showAlert(data.message, 'danger');
            }
        })
        .catch(error => {
            showAlert('上传失败: ' + error, 'danger');
        });
    
    // 清空input以便可以重复选择同一文件
    input.value = '';
}

function deleteKBFile(filename) {
    if (!confirm(`确定要删除文件 "${filename}" 吗？`)) return;
    
    fetch(`/api/kb/delete/${encodeURIComponent(filename)}`, {
        method: 'DELETE'
    })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert(data.message, 'success');
                loadKBFiles();
            } else {
                showAlert(data.message, 'danger');
            }
        });
}

// 更新页面加载逻辑
const originalShowPage = showPage;
showPage = function(page) {
    if (page === 'ai') {
        loadAIPage();
    } else if (page === 'kb-files') {
        loadKBFiles();
    } else if (page === 'logs') {
        loadLogs();
    }
    
    // 隐藏所有页面
    document.querySelectorAll('.page-content').forEach(el => {
        el.style.display = 'none';
    });
    
    // 显示目标页面
    document.getElementById(page + '-page').style.display = 'block';
    
    // 更新菜单激活状态
    document.querySelectorAll('.list-group-item').forEach(el => {
        el.classList.remove('active');
    });
    const menuItem = document.querySelector(`[data-page="${page}"]`);
    if (menuItem) menuItem.classList.add('active');
};
