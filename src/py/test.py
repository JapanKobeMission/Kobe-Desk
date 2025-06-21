import matplotlib.pyplot as plt
plt.figure(figsize=(16, 9))
plt.plot([0, 1], [0, 1])
plt.savefig("test.png", dpi=100)
plt.close()