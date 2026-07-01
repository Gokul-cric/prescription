import numpy as np
import matplotlib.pyplot as plt
import pywt

# Define time series
t = np.linspace(0, 1, 1000, endpoint=False)
x = np.cos(2 * np.pi * 30 * t) + np.sin(2 * np.pi * 60 * t)

# Perform CWT using Morlet wavelet
scales = np.arange(1, 128)
coefficients, frequencies = pywt.cwt(x, scales, 'cmor')

# Plot CWT
plt.figure(figsize=(10, 6))
plt.imshow(np.abs(coefficients), extent=[0, 1, 1, 128], cmap='viridis', aspect='auto',
           vmax=0.5, vmin=0)
plt.colorbar(label='Magnitude')
plt.title('CWT with Morlet Wavelet')
plt.xlabel('Time')
plt.ylabel('Scale')
plt.grid(True, linestyle='--', alpha=0.7)
plt.show()