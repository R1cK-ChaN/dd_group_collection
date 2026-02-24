#!/usr/bin/env bash
# setup-vm-host.sh — Install KVM/QEMU/libvirt and prepare VM resources
# Run on the Linux host as your regular user (uses sudo internally).
set -euo pipefail

ISO_DIR="/var/lib/libvirt/isos"
IMG_DIR="/var/lib/libvirt/images"
VIRTIO_URL="https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/virtio-win.iso"
DISK_PATH="$IMG_DIR/dingtalk-rpa.qcow2"
DISK_SIZE="80G"

echo "=== Step 1: Install virtualization packages ==="
sudo apt update
sudo apt install -y \
    qemu-kvm qemu-utils libvirt-daemon-system libvirt-clients \
    bridge-utils virt-manager ovmf virtinst \
    spice-client-gtk swtpm swtpm-tools

echo "=== Step 2: Add user to libvirt/kvm groups ==="
sudo usermod -aG libvirt,kvm "$USER"

echo "=== Step 3: Enable and start libvirtd ==="
sudo systemctl enable --now libvirtd

echo "=== Step 4: Validate host virtualization ==="
virt-host-validate qemu || true

echo "=== Step 5: Create ISO directory ==="
sudo mkdir -p "$ISO_DIR"

echo "=== Step 6: Download VirtIO drivers ISO ==="
if [ -f "$ISO_DIR/virtio-win.iso" ]; then
    echo "VirtIO ISO already exists, skipping download."
else
    sudo wget -O "$ISO_DIR/virtio-win.iso" "$VIRTIO_URL"
fi

echo "=== Step 7: Create VM disk image ==="
sudo mkdir -p "$IMG_DIR"
if [ -f "$DISK_PATH" ]; then
    echo "Disk image already exists, skipping creation."
else
    sudo qemu-img create -f qcow2 "$DISK_PATH" "$DISK_SIZE"
fi

echo ""
echo "=========================================="
echo "  Host setup complete!"
echo "=========================================="
echo ""
echo "Next steps:"
echo "  1. Place your Windows ISO in $ISO_DIR/"
echo "  2. Log out and back in (for group membership to take effect)"
echo "  3. Run virt-install to create the VM (see README.md)"
echo ""
