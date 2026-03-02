// NextCR Voting System - Additional Features Addon
// Add this script to your index.html after the main script

(function() {
    'use strict';
    
    // Wait for main app to load
    setTimeout(function() {
        
        // ========== FEATURE 1: Export Database ==========
        window.exportDatabase = function() {
            var db = loadDB();
            if (!db) return;
            
            var dataStr = JSON.stringify(db, null, 2);
            var dataBlob = new Blob([dataStr], {type: 'application/json'});
            var url = URL.createObjectURL(dataBlob);
            var link = document.createElement('a');
            link.href = url;
            link.download = 'nextcr_database_' + new Date().toISOString().split('T')[0] + '.json';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
            if (window.toast) {
                window.toast('Database exported!', 'success');
            }
        };
        
        // Add export button to settings if on admin dashboard
        if (window.currentPage === 'admin-dashboard' && window.adminTab === 'settings') {
            var settingsDiv = document.querySelector('.glass');
            if (settingsDiv && !document.getElementById('exportDbBtn')) {
                var exportBtn = document.createElement('button');
                exportBtn.id = 'exportDbBtn';
                exportBtn.className = 'btn btn-primary btn-sm';
                exportBtn.textContent = 'Export Database';
                exportBtn.style.marginTop = '16px';
                exportBtn.onclick = window.exportDatabase;
                
                var lastElement = settingsDiv.lastElementChild;
                if (lastElement) {
                    settingsDiv.insertBefore(exportBtn, lastElement);
                }
            }
        }
        
        // ========== FEATURE 2: Fix Profile Photo Reset ==========
        // Override the photo data reset
        var originalPageProfile = window.pageProfile;
        if (originalPageProfile) {
            window.pageProfile = function(s) {
                // Don't reset photo data if it exists
                if (!window._photoData) {
                    window._photoData = {};
                }
                return originalPageProfile.apply(this, arguments);
            };
        }
        
        var originalPageStudentRegister = window.pageStudentRegister;
        if (originalPageStudentRegister) {
            window.pageStudentRegister = function() {
                // Don't reset photo data if it exists
                if (!window._photoData) {
                    window._photoData = {};
                }
                return originalPageStudentRegister.apply(this, arguments);
            };
        }
        
        // ========== FEATURE 3: Private Chat System ==========
        window.privateChatUsers = {};
        window.currentChatUser = 'all';
        
        // Override sendChat to support private messages
        var originalSendChat = window.sendChat;
        if (originalSendChat) {
            window.sendChat = function() {
                var s = window.getSession();
                if (!s) return;
                
                var inputEl = document.getElementById('chatInput');
                if (!inputEl) return;
                var text = inputEl.value.trim();
                if (!text) return;
                
                var chats = window.loadChats ? window.loadChats() : [];
                var chatMsg = {
                    id: window.uid ? window.uid() : Date.now().toString(),
                    userId: s.user_id,
                    userName: s.name,
                    userRole: s.role,
                    text: text,
                    time: Date.now(),
                    edited: false,
                    deletedForAll: false,
                    deletedFor: [],
                    isPrivate: window.currentChatUser !== 'all',
                    recipient: window.currentChatUser
                };
                
                chats.push(chatMsg);
                if (window.saveChats) {
                    window.saveChats(chats);
                }
                
                inputEl.value = '';
                if (window.navigate) {
                    window.navigate('chat');
                }
            };
        }
        
        // Add user selector to chat
        window.addChatUserSelector = function() {
            var chatArea = document.querySelector('.chat-input-area');
            if (chatArea && !document.getElementById('chatUserSelect')) {
                var db = window.loadDB ? window.loadDB() : null;
                if (!db) return;
                
                var selector = document.createElement('select');
                selector.id = 'chatUserSelect';
                selector.className = 'input-glass';
                selector.style.width = '150px';
                selector.style.marginRight = '8px';
                selector.innerHTML = '<option value="all">Class Chat</option>';
                
                // Add all users as options
                for (var i = 0; i < db.Students.length; i++) {
                    var stu = db.Students[i];
                    selector.innerHTML += '<option value="' + stu.id + '">' + stu.name + '</option>';
                }
                for (var i = 0; i < db.Admins.length; i++) {
                    var adm = db.Admins[i];
                    selector.innerHTML += '<option value="' + adm.id + '">' + adm.name + ' (Admin)</option>';
                }
                
                selector.onchange = function() {
                    window.currentChatUser = this.value;
                    var label = document.getElementById('chatPrivateLabel');
                    if (!label) {
                        label = document.createElement('span');
                        label.id = 'chatPrivateLabel';
                        label.style.color = '#00d4ff';
                        label.style.fontSize = '0.8rem';
                        label.style.marginLeft = '8px';
                        this.parentNode.insertBefore(label, this.nextSibling);
                    }
                    label.textContent = this.value === 'all' ? '' : 'Private Message';
                };
                
                chatArea.insertBefore(selector, chatArea.firstChild);
            }
        };
        
        // Auto-add selector when on chat page
        if (window.currentPage === 'chat') {
            setTimeout(window.addChatUserSelector, 100);
        }
        
        // ========== FEATURE 4: Reset Password Button Style ==========
        // Ensure btn-warning class exists
        if (!document.querySelector('style[data-addon-styles]')) {
            var style = document.createElement('style');
            style.setAttribute('data-addon-styles', 'true');
            style.textContent = `
                .btn-warning {
                    background: linear-gradient(135deg, #ffa502, #ff6348);
                    color: #fff;
                    box-shadow: 0 4px 15px rgba(255,165,2,0.25);
                }
                .btn-warning:hover {
                    transform: translateY(-2px);
                    box-shadow: 0 6px 25px rgba(255,165,2,0.4);
                }
                
                /* Private message styling */
                .chat-msg-private {
                    background: linear-gradient(135deg, rgba(180,74,255,0.1), rgba(255,74,180,0.05)) !important;
                    border-color: rgba(180,74,255,0.2) !important;
                }
                .chat-msg-private::before {
                    content: '🔒 Private';
                    position: absolute;
                    top: 2px;
                    right: 8px;
                    font-size: 0.65rem;
                    color: #b44aff;
                }
                
                /* Message count badge */
                .msg-badge {
                    display: inline-block;
                    background: #ff4757;
                    color: #fff;
                    border-radius: 50%;
                    padding: 2px 6px;
                    font-size: 0.7rem;
                    font-weight: 700;
                    margin-left: 4px;
                    min-width: 18px;
                    text-align: center;
                }
            `;
            document.head.appendChild(style);
        }
        
        // ========== FEATURE 5: Message Count Badge ==========
        window.updateMessageCount = function() {
            var s = window.getSession ? window.getSession() : null;
            if (!s) return;
            
            var chats = window.loadChats ? window.loadChats() : [];
            var unreadCount = 0;
            var lastRead = parseInt(localStorage.getItem('lastReadChat') || '0');
            
            for (var i = 0; i < chats.length; i++) {
                if (chats[i].time > lastRead && chats[i].userId !== s.user_id) {
                    unreadCount++;
                }
            }
            
            // Update chat button in nav
            var chatBtns = document.querySelectorAll('button');
            for (var i = 0; i < chatBtns.length; i++) {
                if (chatBtns[i].textContent.indexOf('Chat') >= 0) {
                    var existingBadge = chatBtns[i].querySelector('.msg-badge');
                    if (existingBadge) {
                        existingBadge.remove();
                    }
                    if (unreadCount > 0) {
                        var badge = document.createElement('span');
                        badge.className = 'msg-badge';
                        badge.textContent = unreadCount;
                        chatBtns[i].appendChild(badge);
                    }
                }
            }
        };
        
        // Update count periodically
        setInterval(window.updateMessageCount, 2000);
        
        // Mark as read when viewing chat
        if (window.currentPage === 'chat') {
            localStorage.setItem('lastReadChat', Date.now());
        }
        
        console.log('[Addon] All features loaded successfully');
        
    }, 500); // Wait for main app to initialize
})();