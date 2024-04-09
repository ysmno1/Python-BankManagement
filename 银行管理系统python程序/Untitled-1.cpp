#include<iostream>
#include<algorithm>
#include<vector>
using namespace std;
vector<int> a, b;
int n;
char op;
int x;
int main() {
    cin>>n;
    while(n--) {
        cin>>op>>x;
        if(op == 'T') a.push_back(x);
        else b.push_back(x);
    }
    b.push_back(1000);
    sort(a.begin(), a.end());
    sort(b.begin(), b.end());
    
    double s = 0;
    double t = 0;
    double v = 1;
    
    int i = 0;
    int j = 0;
    while(i < a.size() || j < b.size()) {
        if(j == b.size() || i < a.size() && a[i] - t < (b[j] - s) * v) {
            s += (a[i] - t) / v;
            t = a[i];
            i += 1;
            v += 1;
        } else {
            t += (b[j] - s) * v;
            s = b[j];
            j += 1;
            v += 1;
        }
    }
    printf("%.0lf", t);
    return 0;
}
