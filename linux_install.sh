#!/bin/bash
set -e  # Exit on any error
# To install it on Linux
# wget https://raw.githubusercontent.com/guimaraeslucas/axonasp/main/linux_install.sh -O linux_install.sh && sed -i 's/\r$//' linux_install.sh && chmod +x linux_install.sh && bash linux_install.sh
# Created by Pieter Cooreman 
# https://github.com/PieterCooreman
# Licensed under the MPL License

echo "========================================="
echo "AxonASP Complete Installation Script"
echo "========================================="

# Ensure we're NOT running as root
if [ "$EUID" -eq 0 ]; then 
    echo "ERROR: Please do NOT run this script as root or with sudo"
    echo "Run it as: ./setup.sh"
    exit 1
fi

# Update system packages
echo "[1/12] Updating system packages..."
sudo yum update -y

# Install Git (if not already installed)
echo "[2/12] Installing Git..."
sudo yum install -y git

# Install Go 1.23.5
echo "[3/12] Installing Go..."
GO_VERSION="1.23.5"
GO_TARBALL="go${GO_VERSION}.linux-amd64.tar.gz"
DOWNLOAD_DIR="$HOME/downloads"
mkdir -p "$DOWNLOAD_DIR"
cd "$DOWNLOAD_DIR"

# Remove old tarball if it exists
rm -f "$GO_TARBALL"

# Download Go
wget -q https://go.dev/dl/${GO_TARBALL}

# Install Go
sudo rm -rf /usr/local/go
sudo tar -C /usr/local -xzf "$GO_TARBALL"

# Cleanup
rm -f "$GO_TARBALL"

# Set up Go environment variables for current user
echo "[4/12] Setting up Go environment..."
if ! grep -q '/usr/local/go/bin' ~/.bashrc; then
    echo 'export PATH=$PATH:/usr/local/go/bin' >> ~/.bashrc
    echo 'export GOPATH=$HOME/go' >> ~/.bashrc
    echo 'export PATH=$PATH:$GOPATH/bin' >> ~/.bashrc
fi
export PATH=$PATH:/usr/local/go/bin
export GOPATH=$HOME/go
export PATH=$PATH:$GOPATH/bin

# Clean up any existing installations
echo "[5/12] Cleaning up old installations..."
if [ -d "$HOME/axonasp" ]; then
    echo "Removing existing installation..."
    rm -rf "$HOME/axonasp"
fi
if sudo test -d "/root/axonasp"; then
    echo "Removing old root installation..."
    sudo rm -rf /root/axonasp
fi

# Stop existing services if running
if sudo systemctl is-active --quiet axonasp 2>/dev/null; then
    echo "Stopping existing AxonASP service..."
    sudo systemctl stop axonasp
fi

# Clone the repository
echo "[6/12] Cloning AxonASP Server repository..."
INSTALL_DIR="$HOME/axonasp"
git clone https://github.com/guimaraeslucas/axonasp.git "$INSTALL_DIR"

# Download Go modules and build
echo "[7/12] Downloading Go modules and building application..."
cd "$INSTALL_DIR"
go mod download
go build -o axonasp

# Ensure proper permissions
chmod +x "$INSTALL_DIR/axonasp"
chmod -R 755 "$INSTALL_DIR"

# Install Nginx
echo "[8/12] Installing Nginx..."
sudo yum install -y nginx

# Get public IP address
echo "[9/12] Detecting public IP address..."
PUBLIC_IP=$(curl -s --max-time 5 http://169.254.169.254/latest/meta-data/public-ipv4 || echo "")

if [ -z "$PUBLIC_IP" ]; then
    echo "Trying alternative method..."
    PUBLIC_IP=$(curl -s --max-time 5 https://api.ipify.org || echo "")
fi

if [ -z "$PUBLIC_IP" ]; then
    SERVER_NAME="_"
    echo "Using wildcard server_name"
else
    echo "Public IP detected: $PUBLIC_IP"
    SERVER_NAME="$PUBLIC_IP"
fi

# Configure Nginx as reverse proxy to port 4050
echo "[10/12] Configuring Nginx..."
sudo tee /etc/nginx/conf.d/axonasp.conf > /dev/null <<EOF
server {
    listen 80 default_server;
    server_name $SERVER_NAME;
    
    location / {
        proxy_pass http://localhost:4050;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_http_version 1.1;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection "upgrade";
    }
}
EOF

# Remove default server configuration to avoid conflicts
if [ -f /etc/nginx/nginx.conf ]; then
    sudo sed -i '/listen.*80 default_server/d' /etc/nginx/nginx.conf
    sudo sed -i '/listen.*\[::\]:80 default_server/d' /etc/nginx/nginx.conf
fi

# Test Nginx configuration
echo "Testing Nginx configuration..."
sudo nginx -t

# Create systemd service for AxonASP
echo "[11/12] Creating systemd service for AxonASP Server..."
sudo tee /etc/systemd/system/axonasp.service > /dev/null <<EOF
[Unit]
Description=AxonASP Server
After=network.target

[Service]
Type=simple
User=$USER
WorkingDirectory=$INSTALL_DIR
ExecStart=$INSTALL_DIR/axonasp
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

# Enable and start services
echo "[12/12] Starting services..."
sudo systemctl daemon-reload

# Enable both services to start on boot
sudo systemctl enable axonasp
sudo systemctl enable nginx

# Start AxonASP service
sudo systemctl start axonasp

# Wait a moment for AxonASP to start
sleep 3

# Start/restart Nginx
sudo systemctl restart nginx

# Configure firewall (if firewalld is running)
if sudo systemctl is-active --quiet firewalld 2>/dev/null; then
    echo "Configuring firewall..."
    sudo firewall-cmd --permanent --add-service=http 2>/dev/null || true
    sudo firewall-cmd --reload 2>/dev/null || true
fi

# Disable SELinux enforcement if it's blocking (common on Amazon Linux)
if command -v getenforce &> /dev/null; then
    if [ "$(getenforce)" == "Enforcing" ]; then
        echo "Setting SELinux to permissive mode..."
        sudo setenforce 0
        # Make it permanent
        sudo sed -i 's/^SELINUX=enforcing/SELINUX=permissive/' /etc/selinux/config 2>/dev/null || true
    fi
fi

# Get display IP
DISPLAY_IP=$PUBLIC_IP
if [ -z "$DISPLAY_IP" ]; then
    DISPLAY_IP=$(curl -s https://api.ipify.org 2>/dev/null || echo "")
fi

# Final status check
sleep 2
AXONASP_STATUS=$(sudo systemctl is-active axonasp)
NGINX_STATUS=$(sudo systemctl is-active nginx)

echo ""
echo "========================================="
echo "âœ“ Installation Complete!"
echo "========================================="
echo "Go version: $(go version)"
echo "AxonASP status: $AXONASP_STATUS"
echo "Nginx status: $NGINX_STATUS"
echo "Installation directory: $INSTALL_DIR"
echo ""
if [ "$DISPLAY_IP" != "" ]; then
    echo "ðŸŒ Website available at: http://$DISPLAY_IP"
else
    echo "ðŸŒ Website available at your instance's public IP"
    echo "   Check EC2 console for your public IP address"
fi
echo ""
echo "Services configured to start automatically on boot."
echo ""
echo "Useful commands:"
echo "  - View AxonASP logs: sudo journalctl -u axonasp -f"
echo "  - View Nginx logs: sudo tail -f /var/log/nginx/access.log"
echo "  - Restart AxonASP: sudo systemctl restart axonasp"
echo "  - Restart Nginx: sudo systemctl restart nginx"
echo "  - Check status: sudo systemctl status axonasp nginx"
echo "========================================="

# Test the installation
echo ""
echo "Testing installation..."
HTTP_CODE=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:4050 2>/dev/null || echo "000")
if [ "$HTTP_CODE" = "200" ] || [ "$HTTP_CODE" = "301" ] || [ "$HTTP_CODE" = "302" ]; then
    echo "âœ“ AxonASP is responding on port 4050 (HTTP $HTTP_CODE)"
else
    echo "âš  Warning: AxonASP may not be responding correctly (HTTP $HTTP_CODE)"
    echo "  Check logs with: sudo journalctl -u axonasp -n 50 --no-pager"
fi

HTTP_CODE=$(curl -s -o /dev/null -w "%{http_code}" http://localhost 2>/dev/null || echo "000")
if [ "$HTTP_CODE" = "200" ] || [ "$HTTP_CODE" = "301" ] || [ "$HTTP_CODE" = "302" ]; then
    echo "âœ“ Nginx is proxying correctly (HTTP $HTTP_CODE)"
else
    echo "âš  Warning: Nginx may not be proxying correctly (HTTP $HTTP_CODE)"
    echo "  Check logs with: sudo tail -50 /var/log/nginx/error.log"
fi

echo ""
echo "Installation script completed successfully!"