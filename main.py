from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import time

def main():
	# Thanks to WoolDoughnut310 for the values: https://github.com/AndreMiras/pycaw/issues/13#issuecomment453862389
	volumes = [64, 56.9, 51.6, 47.7, 44.6, 42, 39.8, 37.8, 36.1, 34.6, 33.2, 31.9, 30.7, 29.6, 28.6, 27.7, 26.8, 25.9, 25.1, 24.3, 23.6, 22.9, 22.2, 21.6, 21, 20.4, 19.8, 19.3, 18.8, 18.3, 17.8, 17.3, 16.8, 16.4, 16, 15.5, 15.1, 14.7, 14.3, 13.9, 13.6, 13.2, 12.9, 12.5, 12.2, 11.8, 11.5, 11.2, 10.9, 10.6, 10.3, 10, 9.7, 9.4, 9.1, 8.9, 8.6, 8.3, 8.1, 7.8, 7.6, 7.3, 7.1, 6.9, 6.6, 6.4, 6.2, 5.9, 5.7, 5.5, 5.3, 5.1, 4.9, 4.7, 4.5, 4.3, 4.1, 3.9, 3.7, 3.5, 3.3, 3.1, 2.9, 2.7, 2.6, 2.4, 2.2, 2, 1.9, 1.7, 1.5, 1.4, 1.2, 1, 0.9, 0.7, 0.6, 0.4, 0.3, 0.1, 0]

	try:
		while True:
			time.sleep(1)
			devices = AudioUtilities.GetSpeakers()
			interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
			volume = cast(interface, POINTER(IAudioEndpointVolume))
			if (volume.GetMute() != 1):
				currentVolume = abs(round(volume.GetMasterVolumeLevel(), 1))
				nearest = 0
				for i in range(1, 100):
					near = abs(currentVolume - volumes[i])
					if (near < abs(currentVolume - volumes[nearest])):
						nearest = i
				nearest = int(round(nearest / 5) * 5)
				volume.SetMasterVolumeLevel(volumes[nearest] * -1, None)
	except:
		# On connection timeout
		time.sleep(1)
		return main()

if __name__ == "__main__":
	main()