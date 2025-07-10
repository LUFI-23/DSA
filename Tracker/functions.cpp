#include <bits/stdc++.h>
using namespace std;

typedef long long ll;

int main() {
    int n , m; cin>>n>>m;
    map<pair<int,int> , ll> mp;

    for(int i = 0;i<m;i++){
        int u , v , w; cin>>u>>v>>w;
        if(mp.find({u, v}) == mp.end()) {
            mp[{u, v}] = w; // Store the weight if the edge is not already present
        } else {
            mp[{u, v}] = min(mp[{u, v}], (ll)w); // Keep the minimum weight if it exists
        }
    }

    vector<vector<pair<int, ll>>> adj(n + 1);
    for(auto& edge : mp) {
        int u = edge.first.first;
        int v = edge.first.second;
        ll w = edge.second;
        adj[u].push_back({v, w});
    }

    vector<vector<ll>> dis(n+1 , vector<ll>(2, LLONG_MAX));

    set<pair<ll,pair<int,int>>> s;
    dis[1][0] = 0; // Distance to the first vertex is 0
    dis[1][1] = 0; // Distance to the first vertex is 0

    s.insert({0, {1, 1}}); // {distance, {vertex, previous_vertex}}

    while(!s.empty()){
        auto it = *s.begin(); 
        s.erase(it);
        ll d = it.first;
        int state = it.second.first;
        int node = it.second.second;
        
        // cout<<node<<" "<<state<<" "<<d<<endl;

        for(auto it:adj[node]){
            int next_node = it.first;
            ll weight = it.second;

            if( d + weight < dis[next_node][state]) {
                s.erase({dis[next_node][state] , {state , next_node}});
                dis[next_node][state] = d + weight; // Update the distance  
                s.insert({dis[next_node][state] , {state , next_node}});
            }
            
            weight = weight / 2; // Halve the weight for the second state
            if( (state == 1) && (d + weight < dis[next_node][0])){
                s.erase({dis[next_node][0], {0, next_node}});
                dis[next_node][0] = d + weight; // Update the distance for the first state
                s.insert({dis[next_node][0], {0, next_node}});
            }
        }
    }
    

    cout<<dis[n][0] << endl; // Output the minimum distance to the nth vertex in the first state
     // Output the distance to the nth vertex
    return 0;
}


// 