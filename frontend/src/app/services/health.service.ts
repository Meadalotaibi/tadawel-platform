import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface HealthResponse {
    status: string;
    timestamp: string;
    uptime: number;
    message: string;
}

@Injectable({
    providedIn: 'root'
})
export class HealthService {
    private apiUrl = 'http://localhost:3000/health';

    constructor(private http: HttpClient) { }

    checkHealth(): Observable<HealthResponse> {
        return this.http.get<HealthResponse>(this.apiUrl);
    }
}
