import numpy as np

# 1 import

# 2
zero_array = np.zeros(10)
print(zero_array)

# 3
ones_array = np.ones(10)
print(ones_array)

# 4
fives_array = np.ones(10)*5
print(fives_array)

# 5
ten_to_fifty_array = np.arange(10,51)
print(ten_to_fifty_array)

# 6
even_ten_to_50 = np.arange(10,51,2)
print(even_ten_to_50)

# 7
matrix_zero_to_eight = np.arange(9).reshape(3,3)
print(matrix_zero_to_eight)

# 8
identity_mat = np.identity(3)
print(identity_mat)

# 9
random_zero_to_one = np.random.rand(1)
print(random_zero_to_one)