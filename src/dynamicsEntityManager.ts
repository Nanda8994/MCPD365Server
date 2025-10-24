import { AuthManager } from './authManager.js';
import Fuse from 'fuse.js';

interface ODataEntity {
    name: string;
    url: string;
}

export class DynamicsEntityManager {
    private entityCache: ODataEntity[] | null = null;
    private authManager = new AuthManager();
    private fuse: Fuse<ODataEntity> | null = null;
    private static readonly FUZZY_THRESHOLD = 0.6;

    public async findBestMatchEntity(query: string): Promise<string | null> {
        if (!this.entityCache) {
            this.entityCache = await this.fetchAllEntities();
            this.fuse = new Fuse(this.entityCache, {
                keys: ['name', 'url'],
                threshold: DynamicsEntityManager.FUZZY_THRESHOLD,
                includeScore: true,
            });
        }
        if (!this.fuse || this.entityCache.length === 0) return null;
        const [result] = this.fuse.search(query);
        return result && result.score !== undefined && result.score <= DynamicsEntityManager.FUZZY_THRESHOLD
            ? result.item.url
            : null;
    }

    private async fetchAllEntities(): Promise<ODataEntity[]> {
        const token = await this.authManager.getToken();
        const url = `${process.env.DYNAMICS_RESOURCE_URL}/data`;
        try {
            const response = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json',
                },
            });
            if (!response.ok) throw new Error(response.statusText);
            const data = await response.json();
            return data.value.map((e: { name: string, url: string }) => ({ name: e.name, url: e.url }));
        } catch {
            return [];
        }
    }
}
