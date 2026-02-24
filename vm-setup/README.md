# Windows VM Setup for DingTalk RPA File Collector

This guide sets up a KVM/QEMU Windows VM on a Linux host to run the DingTalk Group File Collector unattended.

## Prerequisites

- Linux host with KVM support (`/dev/kvm` exists)
- At least 8 GB free RAM and 80 GB free disk
- A Windows ISO (tiny11 or Win10/11 LTSC evaluation)

## Step 1: Host Setup (Automated)

```bash
chmod +x vm-setup/setup-vm-host.sh
./vm-setup/setup-vm-host.sh
```

This installs QEMU/KVM/libvirt, downloads VirtIO drivers, and creates the 80 GB VM disk.

**After running, log out and back in** so the `libvirt`/`kvm` group membership takes effect.

## Step 2: Place Windows ISO

Copy your Windows ISO to the libvirt ISO directory:

```bash
sudo cp /path/to/your-windows.iso /var/lib/libvirt/isos/windows.iso
```

## Step 3: Create the VM

```bash
virt-install \
    --name dingtalk-rpa \
    --ram 8192 --vcpus 4 --cpu host-passthrough \
    --os-variant win10 --boot uefi \
    --tpm backend.type=emulator,backend.version=2.0,model=tpm-crb \
    --disk path=/var/lib/libvirt/images/dingtalk-rpa.qcow2,bus=virtio,cache=writeback \
    --cdrom /var/lib/libvirt/isos/windows.iso \
    --disk path=/var/lib/libvirt/isos/virtio-win.iso,device=cdrom \
    --network network=default,model=virtio \
    --graphics spice,listen=0.0.0.0 --video qxl \
    --channel spicevmc --sound ich9 --noautoconsole
```

> Adjust `--os-variant` to `win11` if using Windows 11. Run `osinfo-query os | grep win` to see options.

## Step 4: Install Windows

1. Open `virt-manager` and connect to the VM
2. During Windows Setup, when asked "Where do you want to install Windows?":
   - Click **Load driver**
   - Browse to the VirtIO CD-ROM: `E:\viostor\w10\amd64` (or `w11`)
   - Also load network driver: `E:\NetKVM\w10\amd64`
3. Complete Windows installation
4. Create a local account (skip Microsoft account)

## Step 5: Install Guest Drivers

Inside the VM:

1. Open the VirtIO CD-ROM (`D:` or `E:`)
2. Run `virtio-win-gt-x64.msi` — installs all VirtIO drivers
3. Install SPICE guest tools from `spice-guest-tools-latest.exe` on the same CD
4. Open **Device Manager** — verify no unknown/missing devices

## Step 6: Run Guest Setup Script

Copy `vm-setup/setup-vm-guest.ps1` into the VM and run as Administrator:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\setup-vm-guest.ps1
```

This installs Python 3.11, pip dependencies, downloads the DingTalk installer, configures auto-login, disables screen lock, and creates a startup shortcut.

## Step 7: Deploy the Project

### Option A: HTTP transfer from host

On the host:
```bash
cd /home/rick/project
python3 -m http.server 8080
```

In the VM browser, download from: `http://192.168.122.1:8080/dd_group_collection/`

Then extract to `C:\dd_group_collection\`.

### Option B: Shared folder

Use SPICE drag-and-drop or set up a virtiofs/9p shared folder.

### Install dependencies

```cmd
cd C:\dd_group_collection
pip install -r requirements.txt
```

## Step 8: Manual Configuration

1. **DingTalk**: Run the installer, log in, enable auto-start in Settings > General
2. **Google Drive for Desktop**: Install, log in — drive letter `G:` will appear
3. **Edit `config.yaml`**: Fill in your actual group names in the `groups` section

## Step 9: Configure VM Auto-start

On the host:
```bash
virsh autostart dingtalk-rpa
```

## Step 10: Test

1. Verify UI automation sees DingTalk:
   ```cmd
   cd C:\dd_group_collection
   python tools\inspect_dingtalk.py
   ```

2. Run one collection cycle:
   ```cmd
   python run.py
   ```

3. **Reboot test**: Restart the VM and verify the full chain:
   - Windows auto-login
   - DingTalk auto-starts
   - Google Drive connects
   - RPA script starts via Startup shortcut

4. **Host reboot test**: Restart the Linux host and verify the VM starts automatically.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| No disk visible during Windows install | Load VirtIO `viostor` driver from CD |
| No network after Windows install | Load VirtIO `NetKVM` driver, or install `virtio-win-gt-x64.msi` |
| VM won't start (permission denied) | Run `sudo chmod 666 /dev/kvm` or check group membership |
| DingTalk UI not found by script | Ensure DingTalk is running and on the main window (not minimized to tray) |
| Google Drive `G:\` not mounted | Open Google Drive app, ensure it's syncing and mounted as `G:` |
