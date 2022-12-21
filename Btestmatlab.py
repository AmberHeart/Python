import numpy as np
import matplotlib.pyplot as plt
import scipy.special as sp

a = 0.5
r = np.linspace(0.1, 3, 234)
theta = np.linspace(0.05, np.pi - 0.05, 200)

[R, T] = np.meshgrid(r, theta)

tmp1 = a / np.sqrt(a**2 + R**2 + 2 * a * R * np.sin(T))
k2 = 4 * a * R * np.sin(T) / (a**2 + R**2 + 2 * a * R * np.sin(T))
tmp2 = (2 - k2) * sp.ellipk(k2) - 2 * sp.ellipe(k2)
tmp3 = tmp1 * tmp2 / k2
result = R * np.sin(T) * tmp3

Y = R * np.cos(T)
X = R * np.sin(T)

plt.contour(X, Y, result, 50)
plt.contour(-X, Y, result, 50)
#plt.contour(X,-Y,T2,30)
plt.axis('equal')
plt.show()