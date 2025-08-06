// Notification System
class NotificationManager {
    constructor() {
        this.notifications = new Map();
        this.idCounter = 0;
        this.container = null;
        this.activeTypes = new Set(); // Track active notification types
        this.init();
    }

    init() {
        // Wait for DOM to be ready
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => {
                this.setupContainer();
            });
        } else {
            this.setupContainer();
        }
    }

    setupContainer() {
        this.container = document.getElementById('notification-container');
        
        // Create container if it doesn't exist
        if (!this.container) {
            this.container = document.createElement('div');
            this.container.id = 'notification-container';
            this.container.className = 'notification-container';
            document.body.insertBefore(this.container, document.body.firstChild);
        }
    }

    show(message, type = 'info', title = '', duration = 5000, options = {}) {
        // Ensure container exists
        if (!this.container) {
            this.setupContainer();
        }
        
        if (!this.container) {
            console.error('Could not create notification container');
            return null;
        }

        // Prevent duplicate notifications of the same type with same message
        const notificationKey = `${type}-${title}-${message.substring(0, 50)}`;
        if (this.activeTypes.has(notificationKey)) {
            console.log('Duplicate notification prevented:', notificationKey);
            return null;
        }

        const id = ++this.idCounter;
        this.activeTypes.add(notificationKey);
        
        // Create notification element
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.dataset.id = id;
        notification.dataset.key = notificationKey;

        // Get icon based on type
        const icons = {
            success: '✅',
            error: '❌',
            warning: '⚠️',
            info: 'ℹ️',
            progress: '🔄'
        };

        const icon = options.icon || icons[type] || icons.info;

        // Create notification content
        notification.innerHTML = `
            <div class="notification-icon">${icon}</div>
            <div class="notification-content">
                ${title ? `<div class="notification-title">${title}</div>` : ''}
                <div class="notification-message">${message}</div>
                ${options.showProgress ? '<div class="notification-progress"><div class="notification-progress-bar"></div></div>' : ''}
            </div>
            <button class="notification-close" onclick="notifications.hide(${id})">×</button>
        `;

        // Add to container
        this.container.appendChild(notification);
        this.notifications.set(id, notification);

        // Trigger show animation
        setTimeout(() => {
            notification.classList.add('show');
        }, 10);

        // Auto hide after duration (unless it's a progress notification)
        if (duration > 0 && type !== 'progress') {
            setTimeout(() => {
                this.hide(id);
            }, duration);
        }

        return id;
    }

    hide(id) {
        const notification = this.notifications.get(id);
        if (!notification) return;

        // Remove from active types
        const key = notification.dataset.key;
        if (key) {
            this.activeTypes.delete(key);
        }

        notification.classList.remove('show');
        notification.classList.add('hide');

        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
            this.notifications.delete(id);
        }, 400);
    }

    updateProgress(id, progress) {
        const notification = this.notifications.get(id);
        if (!notification) return;

        const progressBar = notification.querySelector('.notification-progress-bar');
        if (progressBar) {
            progressBar.style.width = `${progress}%`;
        }

        // Auto hide when progress reaches 100%
        if (progress >= 100) {
            setTimeout(() => {
                this.hide(id);
            }, 1000);
        }
    }

    success(message, title = 'Thành công!', duration = 4000) {
        return this.show(message, 'success', title, duration);
    }

    error(message, title = 'Lỗi!', duration = 7000) {
        return this.show(message, 'error', title, duration);
    }

    warning(message, title = 'Cảnh báo!', duration = 6000) {
        return this.show(message, 'warning', title, duration);
    }

    info(message, title = 'Thông tin', duration = 5000) {
        return this.show(message, 'info', title, duration);
    }

    progress(message, title = 'Đang xử lý...') {
        return this.show(message, 'progress', title, 0, { showProgress: true });
    }

    // Clear all notifications
    clear() {
        this.notifications.forEach((notification, id) => {
            this.hide(id);
        });
    }
}

// Create global instance
const notifications = new NotificationManager();

// Make it available globally
window.notifications = notifications;

// Override the default alert function
window.originalAlert = window.alert;
window.alert = function(message) {
    // Parse message to determine type
    if (message.includes('Lỗi') || message.includes('❌') || message.toLowerCase().includes('error')) {
        notifications.error(message);
    } else if (message.includes('⚠️') || message.includes('Cảnh báo')) {
        notifications.warning(message);
    } else if (message.includes('✅') || message.includes('Thành công')) {
        notifications.success(message);
    } else {
        notifications.info(message);
    }
};