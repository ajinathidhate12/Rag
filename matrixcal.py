def matrixcal(A, B):

    rows_A = len(A)
    cols_A = len(A[0])
    rows_B = len(B)
    cols_B = len(B[0])
    if cols_A != rows_B:
        raise ValueError(f"Cannot multiply: A is {rows_A}x{cols_A}, B is {rows_B}x{cols_B}")
    result = [[0 for _ in range(cols_B)] for _ in range(rows_A)]
    
    for i in range(rows_A):
        for j in range(cols_B):
            for k in range(cols_A):
                result[i][j] += A[i][k] * B[k][j]
    
    return result


def print_matrix(matrix, name=""):
    """Print a matrix in a readable format."""
    if name:
        print(f"{name} =")
    for row in matrix:
        print(f"  {row}")
    print()
A = [
    [1, 2],
    [3, 4]
]

B = [
    [5, 6],
    [7, 8]
]

result = matrixcal(A, B)

print_matrix(A, "A")
print_matrix(B, "B")
print_matrix(result, "A × B")
