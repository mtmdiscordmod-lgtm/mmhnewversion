window.google = {
  script: {
    run: {
      withSuccessHandler: function(callback) {
        return {
          withFailureHandler: function(errCallback) {
             // Return the mock runner that executes the requested function
             return {
                getDataForDashboard: () => callback({
                    authorized: true, email: 'teacher@test.com', courses: []
                }),
                getPendingRequests: (period) => callback([
                    {row: 2, song: 'Mock Song', artist: 'Mock Artist', studentName: 'Student A', videoId: 'dQw4w9WgXcQ'}
                ]),
                getStudentPortalData: () => callback({
                    authorized: true, inventory: 5
                }),
                submitSongRequest: () => callback({success: true, remaining: 4}),
                getLeaderboardData: () => callback({weekly: [], allTime: []}),
                getAuctionData: () => callback({items: [], students: [], periods: ['1']})
             }
          },
          getLeaderboardData: () => callback({weekly: [], allTime: []}),
          getAuctionData: () => callback({items: [], students: [], periods: ['1']}),
          getPendingRequests: (period) => callback([
              {row: 2, song: 'Mock Song', artist: 'Mock Artist', studentName: 'Student A', videoId: 'dQw4w9WgXcQ'}
          ]),
          getStudentPortalData: () => callback({
              authorized: true, inventory: 5
          })
        }
      }
    }
  }
};
