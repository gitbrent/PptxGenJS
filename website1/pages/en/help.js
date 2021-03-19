/**
 * Copyright (c) 2017-present, Brent Ely
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */

const React = require('react');
const CompLibrary = require('../../core/CompLibrary.js');
const Container = CompLibrary.Container;
const GridBlock = CompLibrary.GridBlock;
const siteConfig = require(process.cwd() + '/siteConfig.js');

class Help extends React.Component {
  render() {
    const supportLinks = [
      {
        content: 'Learn more using the [documentation on this site.](/PptxGenJS/docs/installation.html)',
        title: 'Browse Docs',
      },
      {
        content: 'Ask questions about the documentation and project',
        title: 'Join the community',
      },
      {
        content: "Look through the categorized [project issues](https://gitbrent.github.io/PptxGenJS/issues) for a solution",
        title: 'View Issues',
      },
    ];

    return (
      <div className="docMainWrapper wrapper">
        <Container className="mainContainer documentContainer postContainer">
          <div className="post">
            <header className="postHeader">
              <h2>Need help?</h2>
            </header>
            <p>Sometimes implementing a new library can be a difficult task and the slightest mistake will keep something from working. We have all been there!</p>
			<p>If you are having issues getting a presentation to generate, check out the demos in the `examples` directory.</p>
			<p>There are demos for both Nodejs and client-browsers that contain working examples of every available library feature.</p>
			<ul>
				<li>Use a pre-configured jsFiddle to test with: <a href="https://jsfiddle.net/gitbrent/L1uctxm0/">PptxGenJS Fiddle</a></li>
				<li>Use Ask Question on <a href="http://stackoverflow.com/">StackOverflow</a> - be sure to tag it with "PptxGenJS"</li>
			</ul>

            <GridBlock contents={supportLinks} layout="threeColumn" />
          </div>
        </Container>
      </div>
    );
  }
}

module.exports = Help;
