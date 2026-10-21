#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
static int checked_product(size_t count, size_t width, size_t *bytes) {
    if (bytes == NULL || width == 0 || count > SIZE_MAX / width) {
        return 0;
    }
    *bytes = count * width;
    return 1;
}
static int create_zeroed_scores(size_t count, int **out) {
    if (out == NULL || *out != NULL || count == 0) {
        return 0;
    }
    size_t bytes = 0;
    if (!checked_product(count, sizeof **out, &bytes)) {
        return 0;
    }
    int *temporary = calloc(count, sizeof *temporary);
    if (temporary == NULL) {
        return 0;
    }
    *out = temporary;
    return 1;
}
int main(void) {
    int *scores = NULL;
    if (!create_zeroed_scores(5, &scores)) {
        fputs("allocation failed\n", stderr);
        return EXIT_FAILURE;
    }
    int sum = 0;
    for (size_t i = 0; i < 5; ++i) {
        sum += scores[i];
    }
    printf("requested=5 zero_sum=%d\n", sum);
    scores[2] = 90;
    printf("third=%d\n", scores[2]);
    size_t unused = 0;
    int ok = checked_product(SIZE_MAX, 2, &unused);
    printf("overflow_guard=%s\n", ok ? "missed" : "blocked");
    int *zero = NULL;
    ok = create_zeroed_scores(0, &zero);
    printf("zero_request=%s\n", ok ? "accepted" : "rejected");
    free(scores);
    scores = NULL;
    return EXIT_SUCCESS;
}
