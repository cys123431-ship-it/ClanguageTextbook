#include <stdio.h>
#include <stddef.h>
enum { STUDENTS = 3, SUBJECTS = 3 };
int main(void) {
    const int scores[STUDENTS][SUBJECTS] = {
        {90, 80, 70}, {60, 70, 80}, {100, 90, 80}
    };
    for (size_t row = 0; row < STUDENTS; ++row) {
        int sum = 0;
        for (size_t col = 0; col < SUBJECTS; ++col) sum += scores[row][col];
        printf("student %zu: %.2f\n", row + 1, (double)sum / SUBJECTS);
    }
    for (size_t col = 0; col < SUBJECTS; ++col) {
        int sum = 0;
        for (size_t row = 0; row < STUDENTS; ++row) sum += scores[row][col];
        printf("subject %zu: %.2f\n", col + 1, (double)sum / STUDENTS);
    }
    return 0;
}
