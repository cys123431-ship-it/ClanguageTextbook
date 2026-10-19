#include <stdio.h>
#include <stddef.h>

static size_t linear_search(const int *a, size_t n, int key, size_t *steps)
{
    *steps = 0;
    for (size_t i = 0; i < n; ++i) {
        ++*steps;
        if (a[i] == key) return i;
    }
    return n;
}

static size_t binary_search(const int *a, size_t n, int key, size_t *steps)
{
    size_t lo = 0, hi = n;
    *steps = 0;
    while (lo < hi) {
        size_t mid = lo + (hi - lo) / 2;
        ++*steps;
        if (a[mid] == key) return mid;
        if (a[mid] < key) lo = mid + 1;
        else hi = mid;
    }
    return n;
}

int main(void)
{
    const int data[] = {2, 5, 8, 11, 14, 17, 20, 23};
    const int keys[] = {20, 2, 99};
    size_t n = sizeof data / sizeof data[0];
    size_t linear_total = 0, binary_total = 0;
    for (size_t i = 0; i < sizeof keys / sizeof keys[0]; ++i) {
        size_t ls = 0, bs = 0;
        size_t li = linear_search(data, n, keys[i], &ls);
        size_t bi = binary_search(data, n, keys[i], &bs);
        printf("%d: 선형 %zu회, 이진 %zu회, 결과 %s\n",
               keys[i], ls, bs, li == n ? "없음" : "찾음");
        if ((li == n) != (bi == n)) return 1;
        linear_total += ls;
        binary_total += bs;
    }
    printf("비교 총합: 선형 %zu회, 이진 %zu회\n", linear_total, binary_total);
    return 0;
}
