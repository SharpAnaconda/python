# coding: UTF-8
import matplotlib.pyplot as plt

# グラフを描く
x = [1, 2, 3, 4]
y = [1, 4, 9, 16]
plt.plot(x, y, 'bo', label=u"2乗")

# 誤差棒を描く
err = [0.6, 0.7, 2.1, 1.4]
plt.errorbar(x, y, yerr=err, fmt='ro', ecolor='g')

plt.legend()
plt.xlabel(u"X軸")
plt.ylabel(u"Y軸")
plt.title(u"図 1.1")
plt.axis([0, 5, 0, 25])
plt.show()
