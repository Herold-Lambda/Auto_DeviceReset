#include <QApplication>
#include <QMainWindow>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QComboBox>
#include <QPushButton>
#include <QSpinBox>
#include <QLabel>
#include <QTimer>
#include <QMessageBox>
#include <QSystemTrayIcon>
#include <QMenu>
#include <QCloseEvent>
#include <QSettings>
#include <QStyleFactory>
#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <thread>
#include <chrono>
#include <iostream>
#include <string>
#include <vector>
#include <initguid.h>
#include <propkey.h>
#include <Functiondiscoverykeys_devpkey.h>


#ifndef PKEY_Device_FriendlyName
DEFINE_PROPERTYKEY(PKEY_Device_FriendlyName, 0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0, 14);
#endif

DEFINE_GUID(CLSID_MMDeviceEnumerator, 0xBCDE0395, 0xE52F, 0x467C, 0x8E, 0x3D, 0xC4, 0x57, 0x92, 0x91, 0x69, 0x2E);
DEFINE_GUID(IID_IMMDeviceEnumerator, 0xA95664D2, 0x9614, 0x4F35, 0xA7, 0x46, 0xDE, 0x8D, 0xB6, 0x36, 0x17, 0xE6);

// Define the IPolicyConfig interface
interface DECLSPEC_UUID("f8679f50-850a-41cf-9c72-430f290290c8") IPolicyConfig;
class DECLSPEC_UUID("870af99c-171d-4f9e-af0d-e63df40c2bc9") CPolicyConfigClient;

interface IPolicyConfig : public IUnknown
{
public:
    virtual HRESULT GetMixFormat(
            PCWSTR,
            WAVEFORMATEX **
    );

    virtual HRESULT STDMETHODCALLTYPE GetDeviceFormat(
            PCWSTR,
            INT,
            WAVEFORMATEX **
    );

    virtual HRESULT STDMETHODCALLTYPE ResetDeviceFormat(
            PCWSTR
    );

    virtual HRESULT STDMETHODCALLTYPE SetDeviceFormat(
            PCWSTR,
            WAVEFORMATEX *,
            WAVEFORMATEX *
    );

    virtual HRESULT STDMETHODCALLTYPE GetProcessingPeriod(
            PCWSTR,
            INT,
            PINT64,
            PINT64
    );

    virtual HRESULT STDMETHODCALLTYPE SetProcessingPeriod(
            PCWSTR,
            PINT64
    );

    virtual HRESULT STDMETHODCALLTYPE GetShareMode(
            PCWSTR,
            struct DeviceShareMode *
    );

    virtual HRESULT STDMETHODCALLTYPE SetShareMode(
            PCWSTR,
            struct DeviceShareMode *
    );

    virtual HRESULT STDMETHODCALLTYPE GetPropertyValue(
            PCWSTR,
            const PROPERTYKEY &,
            PROPVARIANT *
    );

    virtual HRESULT STDMETHODCALLTYPE SetPropertyValue(
            PCWSTR,
            const PROPERTYKEY &,
            PROPVARIANT *
    );

    virtual HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(
            __in PCWSTR wszDeviceId,
            __in ERole eRole
    );

    virtual HRESULT STDMETHODCALLTYPE SetEndpointVisibility(
            PCWSTR,
            INT
    );
};

struct AudioDevice {
    std::wstring id;
    std::wstring name;
};

std::vector<AudioDevice> GetAudioDevices() {
    std::vector<AudioDevice> devices;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDeviceCollection* pCollection = NULL;

    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        hr = CoCreateInstance(CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
                              IID_IMMDeviceEnumerator, (void**)&pEnumerator);
        if (SUCCEEDED(hr)) {
            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pCollection);
            if (SUCCEEDED(hr)) {
                UINT count;
                pCollection->GetCount(&count);
                for (UINT i = 0; i < count; i++) {
                    IMMDevice* pDevice = NULL;
                    hr = pCollection->Item(i, &pDevice);
                    if (SUCCEEDED(hr)) {
                        LPWSTR pwszID = NULL;
                        hr = pDevice->GetId(&pwszID);
                        if (SUCCEEDED(hr)) {
                            IPropertyStore* pProps = NULL;
                            hr = pDevice->OpenPropertyStore(STGM_READ, &pProps);
                            if (SUCCEEDED(hr)) {
                                PROPVARIANT varName;
                                PropVariantInit(&varName);
                                hr = pProps->GetValue(PKEY_Device_FriendlyName, &varName);
                                if (SUCCEEDED(hr)) {
                                    devices.push_back({pwszID, varName.pwszVal});
                                    PropVariantClear(&varName);
                                }
                                pProps->Release();
                            }
                            CoTaskMemFree(pwszID);
                        }
                        pDevice->Release();
                    }
                }
                pCollection->Release();
            }
            pEnumerator->Release();
        }
        CoUninitialize();
    }
    return devices;
}

HRESULT ToggleAudioDevice(const std::wstring& deviceId, int bufferMs) {
    HRESULT hr = S_OK;
    IPolicyConfig *pPolicyConfig = nullptr;

    try {
        hr = CoInitialize(NULL);
        if (FAILED(hr)) throw std::runtime_error("Failed to initialize COM");

        // Create PolicyConfig instance
        hr = CoCreateInstance(__uuidof(CPolicyConfigClient), NULL, CLSCTX_ALL, __uuidof(IPolicyConfig), (LPVOID *)&pPolicyConfig);
        if (FAILED(hr)) throw std::runtime_error("Failed to create PolicyConfig instance");

        // Deactivate the device
        hr = pPolicyConfig->SetEndpointVisibility(deviceId.c_str(), 0);
        if (FAILED(hr)) throw std::runtime_error("Failed to deactivate the device");

        std::wcout << L"Device " << deviceId << L" has been deactivated" << std::endl;

        // Wait for the specified buffer time
        std::this_thread::sleep_for(std::chrono::milliseconds(bufferMs));

        // Reactivate the device
        hr = pPolicyConfig->SetEndpointVisibility(deviceId.c_str(), 1);
        if (FAILED(hr)) throw std::runtime_error("Failed to reactivate the device");

        std::wcout << L"Device " << deviceId << L" has been reactivated" << std::endl;
    }
    catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
    }

    if (pPolicyConfig) pPolicyConfig->Release();
    CoUninitialize();

    return hr;
}


class AudioDeviceResetGUI : public QMainWindow {
Q_OBJECT

public:
    AudioDeviceResetGUI(bool startMinimized = false) : QMainWindow() {
        setWindowTitle("Audio Device Reset");
        setFixedSize(400, 250);
        setWindowIcon(QIcon(":/icon.ico"));

        QWidget *centralWidget = new QWidget(this);
        setCentralWidget(centralWidget);

        QVBoxLayout *mainLayout = new QVBoxLayout(centralWidget);

        deviceCombo = new QComboBox();
        mainLayout->addWidget(deviceCombo);

        QHBoxLayout *intervalLayout = new QHBoxLayout();
        QLabel *intervalLabel = new QLabel("Reset Interval:");
        hoursSpinBox = new QSpinBox();
        hoursSpinBox->setRange(0, 23);
        hoursSpinBox->setSuffix("h");
        minutesSpinBox = new QSpinBox();
        minutesSpinBox->setRange(0, 59);
        minutesSpinBox->setSuffix("m");
        secondsSpinBox = new QSpinBox();
        secondsSpinBox->setRange(0, 59);
        secondsSpinBox->setSuffix("s");
        intervalLayout->addWidget(intervalLabel);
        intervalLayout->addWidget(hoursSpinBox);
        intervalLayout->addWidget(minutesSpinBox);
        intervalLayout->addWidget(secondsSpinBox);
        mainLayout->addLayout(intervalLayout);

        QHBoxLayout *bufferLayout = new QHBoxLayout();
        QLabel *bufferLabel = new QLabel("Deactivation Buffer:");
        bufferSpinBox = new QSpinBox();
        bufferSpinBox->setRange(1, 10000);
        bufferSpinBox->setSingleStep(100);
        bufferSpinBox->setValue(1000);
        bufferSpinBox->setSuffix("ms");
        bufferLayout->addWidget(bufferLabel);
        bufferLayout->addWidget(bufferSpinBox);
        mainLayout->addLayout(bufferLayout);

        toggleButton = new QPushButton("Toggle Device");
        mainLayout->addWidget(toggleButton);

        startStopButton = new QPushButton("Start");
        mainLayout->addWidget(startStopButton);

        timer = new QTimer(this);

        connect(toggleButton, &QPushButton::clicked, this, &AudioDeviceResetGUI::onToggleClicked);
        connect(startStopButton, &QPushButton::clicked, this, &AudioDeviceResetGUI::onStartStopClicked);
        connect(timer, &QTimer::timeout, this, &AudioDeviceResetGUI::onTimerTimeout);

        populateDevices();
        loadSettings();
        createTrayIcon();

        if (startMinimized) {
            QTimer::singleShot(0, this, &AudioDeviceResetGUI::hide);
        }

        setDarkTheme();
    }

    ~AudioDeviceResetGUI() {
        saveSettings();
    }

    static void setDarkTheme() {
        qApp->setStyle(QStyleFactory::create("Fusion"));

        QPalette darkPalette;
        darkPalette.setColor(QPalette::Window, QColor(53, 53, 53));
        darkPalette.setColor(QPalette::WindowText, Qt::white);
        darkPalette.setColor(QPalette::Base, QColor(25, 25, 25));
        darkPalette.setColor(QPalette::AlternateBase, QColor(53, 53, 53));
        darkPalette.setColor(QPalette::ToolTipBase, Qt::white);
        darkPalette.setColor(QPalette::ToolTipText, Qt::white);
        darkPalette.setColor(QPalette::Text, Qt::white);
        darkPalette.setColor(QPalette::Button, QColor(53, 53, 53));
        darkPalette.setColor(QPalette::ButtonText, Qt::white);
        darkPalette.setColor(QPalette::BrightText, Qt::red);
        darkPalette.setColor(QPalette::Link, QColor(42, 130, 218));
        darkPalette.setColor(QPalette::Highlight, QColor(42, 130, 218));
        darkPalette.setColor(QPalette::HighlightedText, Qt::black);

        qApp->setPalette(darkPalette);

        qApp->setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }");
    }


private slots:
    void onToggleClicked() {
        int index = deviceCombo->currentIndex();
        if (index >= 0 && index < devices.size()) {
            ToggleAudioDevice(devices[index].id, bufferSpinBox->value());
        }
    }

    void onStartStopClicked() {
        if (timer->isActive()) {
            timer->stop();
            startStopButton->setText("Start");
        } else {
            int interval = (hoursSpinBox->value() * 3600 + minutesSpinBox->value() * 60 + secondsSpinBox->value()) * 1000;
            if (interval > 0) {
                timer->start(interval);
                startStopButton->setText("Stop");
            } else {
                QMessageBox::warning(this, "Invalid Interval", "Please set a non-zero interval.");
            }
        }
    }

    void onTimerTimeout() {
        onToggleClicked();
    }

    void trayIconActivated(QSystemTrayIcon::ActivationReason reason) {
        if (reason == QSystemTrayIcon::Trigger) {
            if (isVisible()) {
                hide();
            } else {
                show();
            }
        }
    }

protected:
    void closeEvent(QCloseEvent *event) override {
        if (trayIcon->isVisible()) {
            hide();
            event->ignore();
        }
    }

private:
    void populateDevices() {
        devices = GetAudioDevices();
        for (const auto& device : devices) {
            deviceCombo->addItem(QString::fromStdWString(device.name));
        }
    }

    void saveSettings() {
        QSettings settings("YourCompany", "AudioDeviceReset");
        settings.setValue("selectedDevice", deviceCombo->currentIndex());
        settings.setValue("hours", hoursSpinBox->value());
        settings.setValue("minutes", minutesSpinBox->value());
        settings.setValue("seconds", secondsSpinBox->value());
        settings.setValue("buffer", bufferSpinBox->value());
    }

    void loadSettings() {
        QSettings settings("YourCompany", "AudioDeviceReset");
        int deviceIndex = settings.value("selectedDevice", 0).toInt();
        hoursSpinBox->setValue(settings.value("hours", 0).toInt());
        minutesSpinBox->setValue(settings.value("minutes", 0).toInt());
        secondsSpinBox->setValue(settings.value("seconds", 0).toInt());
        bufferSpinBox->setValue(settings.value("buffer", 1000).toInt());

        // Set the device after populating the combo box
        QTimer::singleShot(0, [this, deviceIndex]() {
            if (deviceIndex >= 0 && deviceIndex < deviceCombo->count()) {
                deviceCombo->setCurrentIndex(deviceIndex);
            }
        });
    }

    void createTrayIcon() {
        trayIconMenu = new QMenu(this);
        trayIconMenu->addAction("Show", this, &QWidget::show);
        trayIconMenu->addAction("Exit", qApp, &QCoreApplication::quit);

        trayIcon = new QSystemTrayIcon(this);
        trayIcon->setContextMenu(trayIconMenu);
        trayIcon->setIcon(QIcon(":/icon.ico"));
        trayIcon->show();

        connect(trayIcon, &QSystemTrayIcon::activated, this, &AudioDeviceResetGUI::trayIconActivated);
    }

    QComboBox *deviceCombo;
    QSpinBox *hoursSpinBox;
    QSpinBox *minutesSpinBox;
    QSpinBox *secondsSpinBox;
    QSpinBox *bufferSpinBox;
    QPushButton *toggleButton;
    QPushButton *startStopButton;
    QTimer *timer;
    std::vector<AudioDevice> devices;
    QSystemTrayIcon *trayIcon;
    QMenu *trayIconMenu;
};

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);

    AudioDeviceResetGUI::setDarkTheme();

    bool startMinimized = false;
    for (int i = 1; i < argc; ++i) {
        if (QString(argv[i]) == "-background") {
            startMinimized = true;
            break;
        }
    }

    AudioDeviceResetGUI gui(startMinimized);
    if (!startMinimized) {
        gui.show();
    }
    return app.exec();
}

#include "main.moc"
