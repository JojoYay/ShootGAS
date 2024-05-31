/**
 * Copyright 2024 JojoYay
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

describe('index', () => {
  describe('payNow', () => {
    it('Returns a hello message', () => {
      expect(generateRemind()).toBe(
        '次回予定6/2(日)リマインドです！\n伝助の更新お忘れなく！\nThis is gentle reminder of 6/2(日).\nPlease update your Densuke schedule.\n〇(14名): 成瀬, Suffian, 安室, Soma, 松平, やまだじょ, 塚本拓, 望月, なべ, おばたけ, 八木, 竹村, 西尾, 西村\n△(4名): SUZUKI, 正田, 德永, 芦田\n伝助URL：https://densuke.biz/list?cd=6MmRJzUzsD3TJPQT'
      );
    });
  });
});
